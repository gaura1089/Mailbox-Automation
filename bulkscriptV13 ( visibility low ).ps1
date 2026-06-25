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

$Users = Import-Csv "$Base\bulk_users.csv"
$OUMap = Get-Content "$Base\OU_Config\OUs.json" -Raw | ConvertFrom-Json
$Cred  = Import-Clixml "$Base\Creds\Cred.xml"

$time = Get-Date -Format "yyyyMMdd_HHmmss"
$SuccessFile = "$Base\Logs\success_$time.csv"
$ErrorFile   = "$Base\Logs\error_$time.csv"

# =========================================
# EXCHANGE CONNECT
# =========================================
$Session = New-PSSession -ConfigurationName Microsoft.Exchange `
 -ConnectionUri http://IN-TZ1-EXMBX2.in.coforgetech.com/PowerShell/ `
 -Authentication Kerberos -Credential $Cred

Import-PSSession $Session -DisableNameChecking -AllowClobber | Out-Null

# =========================================
# PASSWORD
# =========================================
function Generate-RandomPassword {
    $chars="abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%"
    -join ((1..14) | ForEach-Object { $chars[(Get-Random -Max $chars.Length)] })
}

# =========================================
# EMP CHECK (SAFE)
# =========================================
function Check-Emp {
    param($Attr)

    try {
        return Get-ADUser -Filter "Description -eq '$Attr'" -ErrorAction Stop
    }
    catch {
        Write-Host "⚠ Emp check failed (continuing safe)" -ForegroundColor Yellow
        return $null
    }
}

# =========================================
# ✅ FULL ALIAS + SMTP CHECK
# =========================================
function Alias-Exists {
    param($Alias)

    $UPN="$Alias@coforge.com"

    Write-Host "`n🔍 Checking Alias: $Alias" -ForegroundColor Cyan

    # ✅ AD (multi-domain)
    foreach($d in @("IN.COFORGETECH.COM","UK.COFORGETECH.COM","US.COFORGETECH.COM")){
        if(Get-ADUser -Filter "SamAccountName -eq '$Alias'" -Server $d -EA SilentlyContinue){
            Write-Host "   ⚠ Found in AD ($d)"
            return $true
        }
    }

    # ✅ Exchange alias
    if(Get-Recipient -Filter "Alias -eq '$Alias'" -EA SilentlyContinue){
        Write-Host "   ⚠ Found in Exchange Alias"
        return $true
    }

    # ✅ SMTP / ProxyAddresses ✅ CRITICAL
    if(Get-Recipient -Filter "EmailAddresses -like '*$Alias@coforge.com*'" -EA SilentlyContinue){
        Write-Host "   ❌ SMTP exists (ProxyAddresses)"
        return $true
    }

    # ✅ Azure
    try {
        if(Get-MgUser -Filter "userPrincipalName eq '$UPN'" -EA SilentlyContinue){
            Write-Host "   ⚠ Found in Azure"
            return $true
        }
    } catch {}

    Write-Host "   ✅ Alias SAFE"
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

    if(-not(Get-ADUser -Filter "DisplayName -eq '$DN'" -EA SilentlyContinue)){
        return $DN
    }

    $new="$DN - $Emp"

    if(-not(Get-ADUser -Filter "DisplayName -eq '$new'" -EA SilentlyContinue)){
        return $new
    }

    $i=1
    while($true){
        $test="$new-$i"
        if(-not(Get-ADUser -Filter "DisplayName -eq '$test'" -EA SilentlyContinue)){
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

        # ✅ EmpCode format
        $Emp=$u.EmpCode.ToString().PadLeft(8,'0')
        $Attr="$Emp,P"

        Write-Host "`n🔍 Checking EmpCode: $Attr"

        $EmpCheck = Check-Emp $Attr

        if($EmpCheck){

            Write-Host "❌ EmpCode Duplicate: $Attr" -ForegroundColor Red

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
        else{
            Write-Host "✅ EmpCode OK" -ForegroundColor Green
        }

        # ✅ DisplayName
        $DN = if($L){ "$F $L" } else { $F }
        $DN = Get-UniqueDisplayName $DN $Emp

        # ✅ Alias
        $Alias = Get-UniqueAlias $F $L
        $UPN   = "$Alias@coforge.com"

        Write-Host "➡ Processing: $DN ($Alias)"

        # ✅ Check mode (no creation)
        if($IsCheckOnly){
            Write-Host "[CHECK MODE] Preview only" -ForegroundColor Yellow
        }

        # ✅ Dry Run
        elseif($IsDryRun){
            Write-Host "[DRY RUN] Would create mailbox: $Alias" -ForegroundColor Yellow
        }

        # ✅ Actual Execution
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

        # ✅ LOG (ALL MODES)
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
