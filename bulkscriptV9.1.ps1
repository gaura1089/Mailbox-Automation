try {

$ErrorActionPreference = "Continue"
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

$IsDryRun    = $choice -eq "1"
$IsCheckOnly = $choice -eq "3"

# =========================================
# PATHS
# =========================================
$BaseFolder  = "C:\Users\Gaurav.26\OneDrive - Coforge Limited\Documents\UserAutomation"
$OUConfigPath= "$BaseFolder\OU_Config\OUs.json"
$LogFolder   = "$BaseFolder\Logs"
$csvPath     = "$BaseFolder\bulk_users.csv"
$CredPath    = "$BaseFolder\Creds\Cred.xml"

if (!(Test-Path $LogFolder)) {
    New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null
}

$Users = Import-Csv $csvPath
$OUMap = Get-Content $OUConfigPath -Raw | ConvertFrom-Json
$Cred  = Import-Clixml $CredPath

# =========================================
# LOG FILES
# =========================================
$timestamp   = Get-Date -Format "yyyyMMdd_HHmmss"
$SuccessFile = "$LogFolder\bulk_success_$timestamp.csv"
$ErrorFile   = "$LogFolder\bulk_error_$timestamp.csv"

# =========================================
# ✅ EXCHANGE CONNECT
# =========================================
try {
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange `
    -ConnectionUri http://IN-TZ1-EXMBX2.in.coforgetech.com/PowerShell/ `
    -Authentication Kerberos -Credential $Cred

    Import-PSSession $Session -DisableNameChecking -AllowClobber | Out-Null
}
catch {
    Write-Host "⚠ Exchange connection failed: $($_.Exception.Message)" -ForegroundColor Yellow
}

# =========================================
# ✅ PASSWORD
# =========================================
function Generate-RandomPassword {
    $chars="abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%"
    -join ((1..16) | ForEach-Object { $chars[(Get-Random -Maximum $chars.Length)] })
}

# =========================================
# ✅ ALIAS CHECK (ENTERPRISE)
# =========================================
function Alias-Exists {
    param($Alias)

    $UPN = "$Alias@coforge.com"
    $domains = @("IN.COFORGETECH.COM","UK.COFORGETECH.COM","US.COFORGETECH.COM")

    foreach ($domain in $domains){
        if (Get-ADUser -Filter "SamAccountName -eq '$Alias'" -Server $domain -ErrorAction SilentlyContinue){ return $true }
        if (Get-ADUser -Filter "proxyAddresses -like '*$UPN*'" -Properties proxyAddresses -Server $domain -ErrorAction SilentlyContinue){ return $true }
    }

    if (Get-Mailbox -Identity $UPN -ErrorAction SilentlyContinue){ return $true }
    if (Get-RemoteMailbox -Identity $UPN -ErrorAction SilentlyContinue){ return $true }
    if (Get-Recipient -Filter "EmailAddresses -like '*$UPN*'" -ErrorAction SilentlyContinue){ return $true }

    return $false
}

# =========================================
# ✅ ALIAS GENERATOR
# =========================================
function Get-UniqueAlias {
    param($FirstName,$LastName)

    $max = 18
    $FirstName = ($FirstName -replace '\s','').ToLower()
    $LastName  = if([string]::IsNullOrWhiteSpace($LastName)){""}else{($LastName -replace '\s','').ToLower()}

    function free($a){
        if($a.Length -gt $max){ return $false }
        return -not (Alias-Exists $a)
    }

    if($LastName){
        $base="$FirstName.$LastName"
        if(free $base){ return $base }

        $i=1
        while($i -lt 100){
            $new="$FirstName.$i.$LastName"
            if($new.Length -gt $max){ $i++; continue }
            if(free $new){ return $new }
            $i++
        }
    }

    $short = if($LastName){"$FirstName.$($LastName[0])"}else{$FirstName}

    if(free $short){ return $short }

    $i=1
    while($i -lt 100){
        $new = if($LastName){"$FirstName.$i.$($LastName[0])"}else{"$FirstName.$i"}
        if(free $new){ return $new }
        $i++
    }

    throw "Alias generation failed"
}

# =========================================
# ✅ DISPLAY NAME UNIQUE
# =========================================
function Get-UniqueDisplayName {
    param($DN,$Emp)

    if(-not (Get-ADUser -Filter "DisplayName -eq '$DN'" -ErrorAction SilentlyContinue)){
        return $DN
    }

    return "$DN - $Emp"
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

$EmpCodeFormatted = $EmpCode.ToString().PadLeft(8,'0')
$CustomAttr1 = "$EmpCodeFormatted,P"

Write-Host "`n🔍 Checking EmpCode: $CustomAttr1"

# ✅ DUPLICATE EMP
if(Get-ADUser -Filter "Description -eq '$CustomAttr1'" -ErrorAction SilentlyContinue){

    Write-Host "❌ DUPLICATE EMP" -ForegroundColor Red

    [PSCustomObject]@{
        DisplayName="N/A"
        Alias="N/A"
        Email="N/A"
        EmpCode=$EmpCodeFormatted
        License=$License
        Attribute1=$CustomAttr1
        Status="EMP_DUPLICATE"
    } | Export-Csv $ErrorFile -Append -NoTypeInformation

    continue
}

# ✅ EMP CHECK
if($IsCheckOnly){

    Write-Host "✅ EMP CHECK OK → $EmpCodeFormatted"

    [PSCustomObject]@{
        DisplayName="N/A"
        Alias="N/A"
        Email="N/A"
        EmpCode=$EmpCodeFormatted
        License=$License
        Attribute1=$CustomAttr1
        Status="EMP_CHECK_OK"
    } | Export-Csv $SuccessFile -Append -NoTypeInformation

    continue
}

# ✅ DISPLAY NAME
$DisplayNameRaw = if ($user.DisplayName -and $user.DisplayName.Trim() -ne "") { $user.DisplayName } else { "$FirstName $LastName" }
$DisplayName = Get-UniqueDisplayName $DisplayNameRaw $EmpCodeFormatted

# ✅ ALIAS
$Alias = Get-UniqueAlias $FirstName $LastName
$UPN   = "$Alias@coforge.com"

$OUPath= $OUMap.$OUName
if(!$OUPath){ throw "Invalid OU: $OUName" }

# ✅ DRY RUN
if($IsDryRun){

    Write-Host "[DRY RUN] $DisplayName -> $UPN" -ForegroundColor Yellow

    [PSCustomObject]@{
        DisplayName=$DisplayName
        Alias=$Alias
        Email=$UPN
        EmpCode=$EmpCodeFormatted
        License=$License
        Attribute1=$CustomAttr1
        Status="DRY_RUN"
    } | Export-Csv $SuccessFile -Append -NoTypeInformation

    continue
}

# ✅ ACTUAL MAILBOX CREATE ✅🔥
$Password = Generate-RandomPassword
$Secure   = ConvertTo-SecureString $Password -AsPlainText -Force

Write-Host "🚀 Creating → $DisplayName -> $UPN" -ForegroundColor Green

New-RemoteMailbox `
-Name $DisplayName `
-FirstName $FirstName `
-LastName $LastName `
-Alias $Alias `
-UserPrincipalName $UPN `
-OnPremisesOrganizationalUnit $OUPath `
-Password $Secure `
-RemoteRoutingAddress "$Alias@ntlgnoida.mail.onmicrosoft.com" `
-ErrorAction Stop

Start-Sleep 10

Set-RemoteMailbox `
-Identity $UPN `
-CustomAttribute1 $CustomAttr1 `
-CustomAttribute4 $License `
-EmailAddressPolicyEnabled $false `
-ErrorAction Stop

Write-Host "✅ CREATED → $UPN" -ForegroundColor Cyan

[PSCustomObject]@{
    DisplayName=$DisplayName
    Alias=$Alias
    Email=$UPN
    EmpCode=$EmpCodeFormatted
    Password=$Password
    License=$License
    Attribute1=$CustomAttr1
    Status="SUCCESS"
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

# ✅ RESTART
Write-Host "`nPress Y = Restart | Any key = Exit"
$key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

if ($key.Character -eq "Y" -or $key.Character -eq "y") {
    & $PSCommandPath
}
else {
    Read-Host "Press ENTER to exit"
}

}
catch {
    Write-Host "`n🔥 FATAL ERROR 🔥" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Yellow
    Read-Host
}