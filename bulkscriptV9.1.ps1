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

$IsDryRun = $choice -eq "1"
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
$OUMap = Get-Content $OUConfigPath -Raw | ConvertFrom-Json
$Cred = Import-Clixml $CredPath

# =========================================
# LOG FILES
# =========================================
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
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
# PASSWORD
# =========================================
function Generate-RandomPassword {

    $lower   = "abcdefghijklmnopqrstuvwxyz"
    $upper   = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    $numbers = "0123456789"
    $special = "!@#$%"

    # ✅ ensure all categories present
    $passwordChars = @(
        $lower[(Get-Random -Max $lower.Length)]
        $upper[(Get-Random -Max $upper.Length)]
        $numbers[(Get-Random -Max $numbers.Length)]
        $special[(Get-Random -Max $special.Length)]
    )

    # ✅ fill remaining (total 14 chars)
    $all = $lower + $upper + $numbers + $special
    $remaining = 10

    for ($i = 0; $i -lt $remaining; $i++) {
        $passwordChars += $all[(Get-Random -Max $all.Length)]
    }

    # ✅ shuffle password
    return ($passwordChars | Sort-Object {Get-Random}) -join ""
}

# =========================================
# ✅ ALIAS CHECK
# =========================================
function Alias-Exists {
    param($Alias)

    $UPN = "$Alias@coforge.com"

    Write-Host "`n🔍 Checking Alias: $Alias" -ForegroundColor Cyan

    $domains = @("IN.COFORGETECH.COM","UK.COFORGETECH.COM","US.COFORGETECH.COM")

    foreach ($domain in $domains) {
        Write-Host "   ➤ Checking AD ($domain)..."

        # ✅ SamAccountName
        if (Get-ADUser -Filter "SamAccountName -eq '$Alias'" -Server $domain -ErrorAction SilentlyContinue) {
            Write-Host "   ⚠ FOUND in AD SamAccountName ($domain)" -ForegroundColor Yellow
            return $true
        }

        # ✅ ✅ MOST IMPORTANT (ADD THIS ✅🔥)
        if (Get-ADUser -Filter "proxyAddresses -like '*$UPN*'" -Properties proxyAddresses -Server $domain -ErrorAction SilentlyContinue) {
            Write-Host "   ⚠ FOUND in proxyAddresses ($domain)" -ForegroundColor Yellow
            return $true
        }
    }

    if (Get-Mailbox -Identity $UPN -ErrorAction SilentlyContinue) { return $true }
    if (Get-RemoteMailbox -Identity $UPN -ErrorAction SilentlyContinue) { return $true }

    # ✅ better recipient check
    if (Get-Recipient -Filter "EmailAddresses -like '*$UPN*'" -ErrorAction SilentlyContinue) { return $true }

    try {
        if (Get-MgUser -Filter "userPrincipalName eq '$UPN'" -ErrorAction SilentlyContinue) {
            return $true
        }
    } catch {}

    Write-Host "   ✅ NOT FOUND anywhere: $Alias" -ForegroundColor Green
    return $false
}
# =========================================
# ✅ ALIAS GENERATOR
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
        while($i -lt 100){
            $new="$FirstName.$i.$LastName"
            if($new.Length -gt $max){
    $i++
    continue
}
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
    while($i -lt 100){
        $new= if($LastName){"$FirstName.$i.$($LastName[0])"} else {"$FirstName.$i"}
        if($new.Length -le $max){
            if(free $new){return $new}
        }
        $i++
    }
throw "Alias generation failed after 100 attempts"
}


# ✅ DISPLAY NAME UNIQUE
# =========================================
function Get-UniqueDisplayName {
    param($DN, $Emp)

    # ✅ if not exists → return as is
    if(-not (Get-ADUser -Filter "DisplayName -eq '$DN'" -ErrorAction SilentlyContinue)){
        return $DN
    }

    # ✅ if exists → append only EmpCode
    $newName = "$DN - $Emp"

    return $newName
}

# =========================================
# ✅ MAIN LOOP
# =========================================
foreach($user in $Users){

try{

$FirstName = $user.FirstName.Trim()
$LastName  = $user.LastName.Trim()

$EmpCode   = $user.EmpCode
$OUName    = $user.OU
$License   = $user.License


# ✅ EmpCode Check + Log
$EmpCodeFormatted = $EmpCode.ToString().PadLeft(8,'0')
$CustomAttr1 = "$EmpCodeFormatted,P"

# ✅ DisplayName FIX
$DisplayNameRaw = if ($user.DisplayName -and $user.DisplayName.Trim() -ne "") {
    $user.DisplayName.Trim()
} else {
    "$FirstName $LastName"
}

$DisplayName = Get-UniqueDisplayName $DisplayNameRaw $EmpCodeFormatted

Write-Host "`n🔍 Checking EmpCode: $CustomAttr1"

if(Get-ADUser -Filter "Description -eq '$CustomAttr1'" -ErrorAction SilentlyContinue){

Write-Host "❌ DUPLICATE EMP → SKIP" -ForegroundColor Red

[PSCustomObject]@{
DisplayName=$DisplayName
Alias="N/A"
Email="N/A"
EmpCode=$EmpCodeFormatted
Password="N/A"
License=$License
Attribute1=$CustomAttr1
Status="EMP_DUPLICATE"
} | Export-Csv $ErrorFile -Append -NoTypeInformation

continue
}

if($IsCheckOnly){
Write-Host "✅ Emp check mode → skip create"
continue
}

$Alias = Get-UniqueAlias $FirstName $LastName
$UPN   = "$Alias@coforge.com"
$OUPath= $OUMap.$OUName
if(!$OUPath){ throw "Invalid OU: $OUName" }


$Password = Generate-RandomPassword
$Secure   = ConvertTo-SecureString $Password -AsPlainText -Force

Write-Host "➡ Processing: $DisplayName ($Alias)"

if(!$IsDryRun){

New-RemoteMailbox `
-Name $DisplayName `
-FirstName $FirstName `
-LastName $LastName `
-Alias $Alias `
-UserPrincipalName $UPN `
-OnPremisesOrganizationalUnit $OUPath `
-Password $Secure `
-RemoteRoutingAddress "$Alias@ntlgnoida.mail.onmicrosoft.com" `
-ResetPasswordOnNextLogon $false

Start-Sleep 10

Set-RemoteMailbox `
-Identity $UPN `
-CustomAttribute1 $CustomAttr1 `
-CustomAttribute4 $License `
-EmailAddressPolicyEnabled $false

Write-Host "✅ SUCCESS"
}
else{
Write-Host "[DRY RUN] $UPN"
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
Status= if($IsDryRun){"DRY_RUN"}else{"SUCCESS"}
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

Write-Host "`n✅ BULK PROCESS COMPLETED"

# =========================================
# ✅ FIXED END MENU (NO AUTO CLOSE)
# =========================================
Write-Host "`nPress Y = Restart | C = Clear | Any key = Exit"
$key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

if ($key.Character -eq "Y" -or $key.Character -eq "y") {
    & $MyInvocation.MyCommand.Path
}
elseif ($key.Character -eq "C" -or $key.Character -eq "c") {
    Clear-Host
    & $MyInvocation.MyCommand.Path
}
else {
    Write-Host "Exiting..."
    Start-Sleep 2
}