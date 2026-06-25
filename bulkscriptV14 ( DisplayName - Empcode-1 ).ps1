$ErrorActionPreference = "Stop"
Import-Module ActiveDirectory

# =========================================
# ✅ RUN MODE
# =========================================
Write-Host ""
Write-Host "Select Run Mode:" -ForegroundColor Cyan
Write-Host "1. Dry Run"
Write-Host "2. Actual Execution"
Write-Host "3. EmpCode Check Only ✅"

$choice = Read-Host "Enter choice (1/2/3)"

$IsDryRun    = $choice -eq "1"
$IsCheckOnly = $choice -eq "3"

# =========================================
# PATHS
# =========================================
$Base = "C:\Users\Gaurav.26\OneDrive - Coforge Limited\Documents\UserAutomation"

$OUConfigPath = "$Base\OU_Config\OUs.json"
$LogFolder    = "$Base\Logs"
$csvPath      = "$Base\bulk_users.csv"
$CredPath     = "$Base\Creds\Cred.xml"

if (!(Test-Path $LogFolder)) {
    New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null
}

$Users = Import-Csv $csvPath
$OUMap = Get-Content $OUConfigPath -Raw | ConvertFrom-Json
$Cred  = Import-Clixml $CredPath

# =========================================
# ✅ LOG FILES
# =========================================
$timestamp   = Get-Date -Format "yyyyMMdd_HHmmss"
$SuccessFile = "$LogFolder\bulk_success_$timestamp.csv"
$ErrorFile   = "$LogFolder\bulk_error_$timestamp.csv"

# =========================================
# ✅ EXCHANGE CONNECT
# =========================================
$Session = New-PSSession -ConfigurationName Microsoft.Exchange `
 -ConnectionUri http://IN-TZ1-EXMBX2.in.coforgetech.com/PowerShell/ `
 -Authentication Kerberos -Credential $Cred

Import-PSSession $Session -DisableNameChecking -AllowClobber | Out-Null

# =========================================
# ✅ PASSWORD GENERATOR
# =========================================
function Generate-RandomPassword {
    $lower="abcdefghijklmnopqrstuvwxyz"
    $upper="ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    $num="0123456789"
    $sp="!@#$%&*?"

    $p=@()
    $p+=$lower[(Get-Random -Max $lower.Length)]
    $p+=$upper[(Get-Random -Max $upper.Length)]
    $p+=$num[(Get-Random -Max $num.Length)]
    $p+=$sp[(Get-Random -Max $sp.Length)]

    $all=$lower+$upper+$num+$sp

    for($i=0;$i -lt 10;$i++){
        $p+=$all[(Get-Random -Max $all.Length)]
    }

    return (-join ($p | Sort-Object {Get-Random}))
}

# =========================================
# ✅ EMP CHECK FUNCTION
# =========================================
function Check-Emp {
    param($Attr)

    try {
        return Get-ADUser -Filter "Description -eq '$Attr'" -ErrorAction Stop
    }
    catch {
        Write-Host "⚠ Emp check failed (AD issue)" -ForegroundColor Yellow
        return $null
    }
}

# =========================================
# ✅ ALIAS CHECK (FULL VISIBILITY)
# =========================================
function Alias-Exists {
    param($Alias)

    $UPN="$Alias@coforge.com"

    Write-Host "`n🔍 Checking Alias: $Alias" -ForegroundColor Cyan

    # ✅ AD CHECK
    foreach($d in @("IN.COFORGETECH.COM","UK.COFORGETECH.COM","US.COFORGETECH.COM")){
        Write-Host "   ➤ Checking AD ($d)..."
        if(Get-ADUser -Filter "SamAccountName -eq '$Alias'" -Server $d -ErrorAction SilentlyContinue){
            Write-Host "   ⚠ FOUND in AD ($d)" -ForegroundColor Yellow
            return $true
        }
    }

    # ✅ MAILBOX
    Write-Host "   ➤ Checking Exchange Mailbox..."
    if(Get-Mailbox -Identity $UPN -ErrorAction SilentlyContinue){
        Write-Host "   ⚠ FOUND in Mailbox" -ForegroundColor Yellow
        return $true
    }

    # ✅ REMOTE MAILBOX
    Write-Host "   ➤ Checking RemoteMailbox..."
    if(Get-RemoteMailbox -Identity $UPN -ErrorAction SilentlyContinue){
        Write-Host "   ⚠ FOUND in RemoteMailbox" -ForegroundColor Yellow
        return $true
    }

    # ✅ RECIPIENT
    Write-Host "   ➤ Checking Exchange Alias..."
    if(Get-Recipient -Filter "Alias -eq '$Alias'" -ErrorAction SilentlyContinue){
        Write-Host "   ⚠ FOUND in Recipient" -ForegroundColor Yellow
        return $true
    }

    # ✅ SMTP / PROXY
    Write-Host "   ➤ Checking SMTP / ProxyAddresses..."
    if(Get-Recipient -Filter "EmailAddresses -like '*$UPN*'" -ErrorAction SilentlyContinue){
        Write-Host "   ❌ SMTP EXISTS" -ForegroundColor Red
        return $true
    }

    # ✅ AZURE
    Write-Host "   ➤ Checking Azure AD..."
    try {
        if(Get-MgUser -Filter "userPrincipalName eq '$UPN'" -ErrorAction SilentlyContinue){
            Write-Host "   ⚠ FOUND in Azure" -ForegroundColor Yellow
            return $true
        }
    } catch {}

    Write-Host "   ✅ Alias SAFE" -ForegroundColor Green
    return $false
}

# =========================================
# ✅ ALIAS GENERATOR
# =========================================
function Get-UniqueAlias {
    param($F,$L)

    $F=($F -replace '\s','').ToLower()
    $L= if([string]::IsNullOrWhiteSpace($L)) {""} else {($L -replace '\s','').ToLower()}

    $base="$F.$L"

    if(-not (Alias-Exists $base)){ return $base }

    $i=1
    while($true){
        $new="$F.$i.$L"
        if(-not (Alias-Exists $new)){ return $new }
        $i++
    }
}

# =========================================
# ✅ DISPLAY NAME UNIQUE
# =========================================
function Get-UniqueDisplayName {
    param($DN,$Emp)

    if(-not(Get-ADUser -Filter "DisplayName -eq '$DN'" -ErrorAction SilentlyContinue)){
        return $DN
    }

    $new="$DN - $Emp"

    $i=1
    while($true){
        $test="$new-$i"
        if(-not(Get-ADUser -Filter "DisplayName -eq '$test'" -ErrorAction SilentlyContinue)){
            return $test
        }
        $i++
    }
}

# =========================================
# ✅ MAIN LOOP
# =========================================
foreach($u in $Users){

    $F="";$L="";$Alias="";$UPN="";$Pass="";$Emp="";$Attr="";$DN=""

    try{

        $F=$u.FirstName.Trim()
        $L=$u.LastName.Trim()

        if([string]::IsNullOrWhiteSpace($F)){
            throw "FirstName missing"
        }

        # ✅ EMP FORMAT
        $Emp=$u.EmpCode.ToString().PadLeft(8,'0')
        $Attr="$Emp,P"

        Write-Host "`n🔍 Checking EmpCode: $Attr"

        $EmpCheck = Check-Emp $Attr

        if($EmpCheck){

            Write-Host "❌ EmpCode Duplicate" -ForegroundColor Red

            [PSCustomObject]@{
                DisplayName="$F $L"
                Alias="N/A"
                Email="N/A"
                EmpCode=$Emp
                Password="N/A"
                License=$u.License
                Attribute1=$Attr
                Status="EMP_DUPLICATE"
            } | Export-Csv $ErrorFile -Append -NoTypeInformation

            continue
        }

        Write-Host "✅ EmpCode OK" -ForegroundColor Green

        # ✅ DISPLAY NAME
        $DN = if($L){ "$F $L" } else { $F }
        $DN = Get-UniqueDisplayName $DN $Emp

        # ✅ ALIAS
        $Alias = Get-UniqueAlias $F $L
        $UPN   = "$Alias@coforge.com"

        Write-Host "➡ Processing: $DN ($Alias)" -ForegroundColor Cyan

        # ✅ MODE HANDLING
        if($IsCheckOnly){
            Write-Host "[CHECK MODE] Preview only" -ForegroundColor Yellow
        }
        elseif($IsDryRun){
            Write-Host "[DRY RUN] Would create mailbox: $Alias" -ForegroundColor Yellow
        }
        else{

            $OUPath = $OUMap.$($u.OU)
            if(!$OUPath){ throw "Invalid OU: $($u.OU)" }

            $Pass = Generate-RandomPassword

            New-RemoteMailbox `
                -Name $DN `
                -FirstName $F `
                -LastName $L `
                -Alias $Alias `
                -UserPrincipalName $UPN `
                -OnPremisesOrganizationalUnit $OUPath `
                -Password (ConvertTo-SecureString $Pass -AsPlainText -Force) `
                -RemoteRoutingAddress "$Alias@ntlgnoida.mail.onmicrosoft.com"

            Start-Sleep 5

            Set-RemoteMailbox `
                -Identity $UPN `
                -CustomAttribute1 $Attr `
                -CustomAttribute4 $u.License `
                -EmailAddressPolicyEnabled $false

            Write-Host "✅ CREATED" -ForegroundColor Green
        }

        # ✅ LOG
        [PSCustomObject]@{
            DisplayName=$DN
            Alias=$Alias
            Email=$UPN
            EmpCode=$Emp
            Password=$Pass
            License=$u.License
            Attribute1=$Attr
            Status= if($IsCheckOnly){"CHECK_ONLY"} elseif($IsDryRun){"DRY_RUN"} else {"SUCCESS"}
        } | Export-Csv $SuccessFile -Append -NoTypeInformation

    }
    catch{

        Write-Host "❌ ERROR: $($_.Exception.Message)" -ForegroundColor Red

        [PSCustomObject]@{
            DisplayName="$($u.FirstName) $($u.LastName)"
            Alias=$Alias
            Email=$UPN
            EmpCode=$u.EmpCode
            Password=$Pass
            License=$u.License
            Attribute1=$Attr
            Status="FAILED"
        } | Export-Csv $ErrorFile -Append -NoTypeInformation
    }
}

Write-Host "`n✅ BULK COMPLETED ✅" -ForegroundColor Green

# =========================================
# ✅ RESTART
# =========================================
$resp = Read-Host "Press Y to restart or any key to exit"

if($resp -match "y"){
    Write-Host "🔁 Restarting..."
    Start-Sleep 1
    & $MyInvocation.MyCommand.Path
}
else{
    Write-Host "⏹ Exit"
}

