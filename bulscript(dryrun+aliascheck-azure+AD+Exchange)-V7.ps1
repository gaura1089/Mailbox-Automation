$ErrorActionPreference = "Stop"
Import-Module ActiveDirectory

# =========================================
# DRY RUN SELECTION ✅
# =========================================
Write-Host ""
Write-Host "Select Run Mode:" -ForegroundColor Cyan
Write-Host "1. Dry Run (No changes)"
Write-Host "2. Actual Execution"
$choice = Read-Host "Enter choice (1/2)"

$IsDryRun = $choice -eq "1"

if ($IsDryRun) {
    Write-Host "✅ DRY RUN MODE ENABLED" -ForegroundColor Yellow
} else {
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
    $chars="abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%"
    -join ((1..12) | ForEach-Object { $chars[(Get-Random -Maximum $chars.Length)] })
}

# =========================================
# DOMAIN + EXCHANGE + SMTP + AZURE ✅🔥
# =========================================
function Alias-Exists {
    param($Alias)

    $UPN = "$Alias@coforge.com"
    $UPNLower = $UPN.ToLower()

    $domains = @("IN.COFORGETECH.COM","UK.COFORGETECH.COM","US.COFORGETECH.COM")

    # ✅ AD CHECK
    foreach ($domain in $domains) {
        if (Get-ADUser -Filter "SamAccountName -eq '$Alias'" -Server $domain -ErrorAction SilentlyContinue) {
            Write-Host "⚠ Found in AD: $Alias ($domain)" -ForegroundColor Yellow
            return $true
        }
    }

    # ✅ EXCHANGE FULL CHECK (STRICT)
    $recipients = Get-Recipient -ResultSize Unlimited -ErrorAction SilentlyContinue

    foreach ($r in $recipients) {

        # Alias match
        if ($r.Alias -eq $Alias) {
            Write-Host "⚠ Found in Exchange Alias: $Alias" -ForegroundColor Yellow
            return $true
        }

        # UPN match
        if ($r.UserPrincipalName -and ($r.UserPrincipalName.ToLower() -eq $UPNLower)) {
            Write-Host "⚠ Found Exchange UPN: $UPN" -ForegroundColor Yellow
            return $true
        }

        # SMTP match (MOST RELIABLE 🔥)
        foreach ($mail in $r.EmailAddresses) {
            if ($mail.ToString().ToLower() -like "*$UPNLower*") {
                Write-Host "⚠ Found Exchange SMTP: $UPN" -ForegroundColor Yellow
                return $true
            }
        }
    }

    # ✅ AZURE FULL CHECK (STRICT ✅🔥)
    try {
        $allUsers = Get-MgUser -All -ErrorAction SilentlyContinue

        foreach ($user in $allUsers) {
            if ($user.UserPrincipalName -and ($user.UserPrincipalName.ToLower() -eq $UPNLower)) {
                Write-Host "⚠ Found in Azure AD: $UPN" -ForegroundColor Yellow
                return $true
            }
        }
    }
    catch {
        Write-Host "⚠ Azure check failed (ignored)" -ForegroundColor DarkYellow
    }

    return $false
}

# =========================================
# FINAL ALIAS LOGIC ✅
# =========================================
function Get-UniqueAlias {
    param($FirstName,$LastName)

    $max=30-12

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

    $short= if($LastName){"$FirstName.$($LastName[0])"} else {$FirstName}

    if($short.Length -gt $max){
        $short=$FirstName.Substring(0,$max)
    }

    if(free $short){return $short}

    $i=1
    while($true){
        $new= if($LastName){"$FirstName.$i.$($LastName[0])"} else {"$FirstName.$i"}
        if($new.Length -le $max){
            if(free $new){return $new}
        }
        $i++
    }
}

# =========================================
# MAIN LOOP
# =========================================
foreach($user in $Users){

    try {

        $FirstName=$user.FirstName
        $LastName=$user.LastName
        $EmpCode=$user.EmpCode
        $OUName=$user.OU
        $License=$user.License

        if([string]::IsNullOrWhiteSpace($FirstName)){
            throw "FirstName missing"
        }

        $OUPath=$OUMap.$OUName
        if(!$OUPath){ throw "Invalid OU: $OUName" }

        if([string]::IsNullOrWhiteSpace($LastName)){
            $DisplayName=$FirstName
        } else {
            $DisplayName="$FirstName $LastName"
        }

        if(Get-ADUser -Filter "DisplayName -eq '$DisplayName'" -ErrorAction SilentlyContinue){
            $DisplayName="$DisplayName - $EmpCode"
        }

        $Alias=Get-UniqueAlias $FirstName $LastName
        $UPN="$Alias@coforge.com"
        $Routing="$Alias@ntlgnoida.mail.onmicrosoft.com"

        $Password=Generate-RandomPassword
        $SecurePassword=ConvertTo-SecureString $Password -AsPlainText -Force

        Write-Host "➡ Processing: $DisplayName ($Alias)"

        if(!$IsDryRun){

            New-RemoteMailbox `
                -Name $DisplayName `
                -FirstName $FirstName `
                -LastName $LastName `
                -Alias $Alias `
                -UserPrincipalName $UPN `
                -OnPremisesOrganizationalUnit $OUPath `
                -Password $SecurePassword `
                -RemoteRoutingAddress $Routing `
                -ResetPasswordOnNextLogon $false `
                -ErrorAction Stop

            Start-Sleep 10

            $CustomAttr1="$EmpCode,P"

            Set-RemoteMailbox `
                -Identity $UPN `
                -CustomAttribute1 $CustomAttr1 `
                -CustomAttribute4 $License `
                -EmailAddressPolicyEnabled $false `
                -ErrorAction Stop

            Write-Host "✅ SUCCESS: $DisplayName" -ForegroundColor Green
        }
        else {
            Write-Host "[DRY RUN] Would create mailbox: $Alias" -ForegroundColor Yellow
        }

        # ✅ ALWAYS SET FOR LOG
        $CustomAttr1="$EmpCode,P"

        [PSCustomObject]@{
            DisplayName = $DisplayName
            Alias       = $Alias
            Email       = $UPN
            EmpCode     = $EmpCode
            Password    = $Password
            License     = $License
            Attribute1  = $CustomAttr1
            Status      = if($IsDryRun){"DRY_RUN"} else {"SUCCESS"}
        } | Export-Csv $SuccessFile -Append -NoTypeInformation

    }
    catch{
        Write-Host "❌ ERROR: $($_.Exception.Message)" -ForegroundColor Red

        [PSCustomObject]@{
            DisplayName="$($user.FirstName) $($user.LastName)"
            EmpCode=$user.EmpCode
            Error=$_.Exception.Message
            Status="FAILED"
        } | Export-Csv $ErrorFile -Append -NoTypeInformation
    }
}

Write-Host ""
Write-Host "✅ BULK PROCESS COMPLETED ✅" -ForegroundColor Green

Write-Host "Press Y within 10 seconds to run again OR wait to auto exit..." -ForegroundColor Cyan

$startTime = Get-Date
$inputReceived = $false
$choice = ""

while ((Get-Date) -lt $startTime.AddSeconds(10)) {

    if ($Host.UI.RawUI.KeyAvailable) {
        $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        $choice = $key.Character
        $inputReceived = $true
        break
    }

    Start-Sleep -Milliseconds 200
}

if ($inputReceived -and ($choice -eq "Y" -or $choice -eq "y")) {

    Write-Host "🔁 Restarting Script..."
    Start-Sleep 1
    & $MyInvocation.MyCommand.Path  # ✅ FIXED
}
else {
    Write-Host "⏹ Auto exiting..."
    Start-Sleep 3
}
