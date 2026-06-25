$ErrorActionPreference = "Stop"
Import-Module ActiveDirectory

# =========================================
# ✅ RUN MODE
# =========================================
Write-Host ""
Write-Host "Select Run Mode:" -ForegroundColor Cyan
Write-Host "1. Dry Run (No changes)"
Write-Host "2. Actual Execution"
Write-Host "3. EmpCode Check Only ✅"

$choice = Read-Host "Enter choice (1/2/3)"

$IsDryRun   = $choice -eq "1"
$IsCheckOnly = $choice -eq "3"

if ($IsDryRun) {
    Write-Host "✅ DRY RUN MODE ENABLED" -ForegroundColor Yellow
}
elseif ($IsCheckOnly) {
    Write-Host "✅ EMP CODE CHECK ONLY MODE ENABLED" -ForegroundColor Cyan
}
else {
    Write-Host "✅ ACTUAL EXECUTION MODE" -ForegroundColor Green
}

# =========================================
# PATHS
# =========================================
$BaseFolder = "C:\Users\Gaurav.26\OneDrive - Coforge Limited\Documents\UserAutomation"
$OUConfigPath = "$BaseFolder\OU_Config\OUs.json"
$LogFolder = "$BaseFolder\Logs"
$csvPath = "$BaseFolder\bulk_users.csv"
$CredPath = "$BaseFolder\Creds\Cred.xml"

if (!(Test-Path $LogFolder)) {
    New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null
}

$Users = Import-Csv $csvPath

# ✅ CSV VALIDATION
if (!$Users[0].PSObject.Properties.Name -contains "FirstName" -or
    !$Users[0].PSObject.Properties.Name -contains "LastName") {
    throw "❌ CSV must contain FirstName & LastName"
}

$OUMap = Get-Content $OUConfigPath -Raw | ConvertFrom-Json
$Cred = Import-Clixml $CredPath

# ✅ LOG FILES
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$SuccessFile = "$LogFolder\bulk_success_$timestamp.csv"
$ErrorFile   = "$LogFolder\bulk_error_$timestamp.csv"

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

    $lower = "abcdefghijklmnopqrstuvwxyz"
    $upper = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    $numbers = "0123456789"
    $special = "!@#$%&*?"

    $password = @()
    $password += $lower[(Get-Random -Maximum $lower.Length)]
    $password += $upper[(Get-Random -Maximum $upper.Length)]
    $password += $numbers[(Get-Random -Maximum $numbers.Length)]
    $password += $special[(Get-Random -Maximum $special.Length)]

    $all = $lower + $upper + $numbers + $special
    $remainingLength = Get-Random -Minimum 10 -Maximum 12

    for ($i=0; $i -lt $remainingLength; $i++) {
        $password += $all[(Get-Random -Maximum $all.Length)]
    }

    return (-join ($password | Sort-Object {Get-Random}))
}

# =========================================
# ✅ ALIAS CHECK
# =========================================
function Alias-Exists {
    param($Alias)

    $UPN = "$Alias@coforge.com"
    $domains = @("IN.COFORGETECH.COM","UK.COFORGETECH.COM","US.COFORGETECH.COM")

    foreach ($domain in $domains) {
        if (Get-ADUser -Filter "SamAccountName -eq '$Alias'" -Server $domain -ErrorAction SilentlyContinue) {
            return $true
        }
    }

    if (Get-Mailbox -Identity $UPN -ErrorAction SilentlyContinue) { return $true }
    if (Get-RemoteMailbox -Identity $UPN -ErrorAction SilentlyContinue) { return $true }
    if (Get-Recipient -Filter "Alias -eq '$Alias'" -ErrorAction SilentlyContinue) { return $true }

    try {
        if (Get-MgUser -Filter "userPrincipalName eq '$UPN'" -ErrorAction SilentlyContinue) {
            return $true
        }
    } catch {}

    return $false
}

# =========================================
# ✅ ALIAS GENERATION
# =========================================
function Get-UniqueAlias {
    param($FirstName,$LastName)

    $max=18

    $FirstName=($FirstName -replace '\s','').ToLower()
    $LastName= if([string]::IsNullOrWhiteSpace($LastName)) {""} else {($LastName -replace '\s','').ToLower()}

    function free($a){
        if($a.Length -gt $max){return $false}
        return -not(Alias-Exists $a)
    }

    if($LastName){
        $full="$FirstName.$LastName"
        if(free $full){return $full}

        $i=1
        while($true){
            $new="$FirstName.$i.$LastName"
            if($new.Length -gt $max){break}
            if(free $new){return $new}
            $i++
        }
    }

    return $FirstName
}

# =========================================
# ✅ MAIN LOOP
# =========================================
foreach($user in $Users){

    $FirstName=""; $LastName=""; $Alias=""; $UPN=""; $Password=""
    $EmpCodeFormatted=""; $CustomAttr1=""; $DisplayName=""
    $License=$user.License

    try {

        $FirstName = $user.FirstName.Trim()
        $LastName  = $user.LastName.Trim()

        if([string]::IsNullOrWhiteSpace($FirstName)){
            throw "FirstName missing"
        }

        $EmpCode=$user.EmpCode
        $OUName=$user.OU

        $EmpCodeFormatted = $EmpCode.ToString().Trim().PadLeft(8,'0')
        $CustomAttr1 = "$EmpCodeFormatted,P"

        $EmpCheck = Get-ADUser -Filter "Description -eq '$CustomAttr1'" -ErrorAction SilentlyContinue

        if ($EmpCheck) {
            Write-Host "❌ Duplicate EmpCode"
            continue
        }

        if ($IsCheckOnly) {
            Write-Host "✅ EmpCode OK"
            continue
        }

        $OUPath=$OUMap.$OUName

        # ✅ DisplayName FIX
        $DisplayName="$FirstName $LastName"

        if (Get-ADUser -Filter "DisplayName -eq '$DisplayName'" -ErrorAction SilentlyContinue) {
            $DisplayName="$DisplayName - $EmpCodeFormatted"
        }

        $Alias=Get-UniqueAlias $FirstName $LastName
        $UPN="$Alias@coforge.com"

        $Password=Generate-RandomPassword
        $SecurePassword=ConvertTo-SecureString $Password -AsPlainText -Force

        if(!$IsDryRun){

            New-RemoteMailbox `
                -Name $DisplayName `
                -FirstName $FirstName `
                -LastName $LastName `
                -Alias $Alias `
                -UserPrincipalName $UPN `
                -OnPremisesOrganizationalUnit $OUPath `
                -Password $SecurePassword `
                -RemoteRoutingAddress "$Alias@ntlgnoida.mail.onmicrosoft.com"

        }

        # ✅ SUCCESS LOG
        [PSCustomObject]@{
            DisplayName=$DisplayName
            Alias=$Alias
            Email=$UPN
            EmpCode=$EmpCodeFormatted
            Password=$Password
            License=$License
            Attribute1=$CustomAttr1
            Status= if($IsDryRun){"DRY_RUN"} else {"SUCCESS"}
        } | Export-Csv $SuccessFile -Append -NoTypeInformation

    }
    catch{
        Write-Host "❌ ERROR: $($_.Exception.Message)"

        [PSCustomObject]@{
            DisplayName="$FirstName $LastName"
            Alias=$Alias
            Email=$UPN
            EmpCode=$EmpCodeFormatted
            Password=$Password
            License=$License
            Attribute1=$CustomAttr1
            Status="FAILED"
        } | Export-Csv $ErrorFile -Append -NoTypeInformation
    }
}

Write-Host "`n✅ COMPLETED ✅"

# =========================================
# ✅ RESTART
# =========================================
Write-Host "Press Y to rerun..."

$key = Read-Host

if ($key -eq "Y" -or $key -eq "y") {
    & $MyInvocation.MyCommand.Path
}