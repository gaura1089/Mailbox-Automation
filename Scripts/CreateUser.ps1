param(
    [string]$Mode,
    [string]$FirstName,
    [string]$LastName,
    [string]$DisplayName,
    [string]$OU,
    [string]$EmpCode,
    [string]$License,
    [string]$HRName
)

Import-Module ActiveDirectory

# =========================================
# PATHS
# =========================================
$BaseFolder = "C:\Users\Gaurav.26\OneDrive - Coforge Limited\Documents\UserAutomation"
$LogFolder = "$BaseFolder\Logs"
$OUConfigPath = "$BaseFolder\OU_Config\OUs.json"

if (!(Test-Path $LogFolder)) {
    New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null
}

if ($Mode -ne "Single") {
    Write-Host "INVALID MODE" -ForegroundColor Red
    exit 1
}

# =========================================
# OU MAPPING
# =========================================
try {
    $OUMap = Get-Content $OUConfigPath -Raw | ConvertFrom-Json
} catch {
    Write-Host "OU JSON not found ❌"
    exit 1
}

$OUPath = $OUMap.$OU

if (-not $OUPath) {

    $ErrorLog = "$LogFolder\$EmpCode`_ERROR.csv"

    [PSCustomObject]@{
        EmpCode=$EmpCode
        DisplayName=$DisplayName
        Status="FAILED"
        Error="INVALID OU : $OU"
    } | Export-Csv $ErrorLog -NoTypeInformation -Encoding UTF8

    exit 1
}

# =========================================
# EXCHANGE CONNECT
# =========================================
$Cred = Import-Clixml "$BaseFolder\Creds\Cred.xml"

$Session = New-PSSession -ConfigurationName Microsoft.Exchange `
    -ConnectionUri http://IN-TZ1-EXMBX2.in.coforgetech.com/PowerShell/ `
    -Authentication Kerberos -Credential $Cred

Import-PSSession $Session -DisableNameChecking -AllowClobber | Out-Null

# =========================================
# PASSWORD
# =========================================
function Generate-RandomPassword {
    param([int]$Length = 15)
    $chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()_-+=?"
    -join ((1..$Length) | ForEach-Object {
        $chars[(Get-Random -Maximum $chars.Length)]
    })
}

# =========================================
# DOMAIN CHECK
# =========================================
function Alias-Exists {
    param($Alias)

    $domains=@("IN.COFORGETECH.COM","UK.COFORGETECH.COM","US.COFORGETECH.COM")

    foreach($domain in $domains){
        $check = Get-ADUser -Filter "SamAccountName -eq '$Alias'" -Server $domain -ErrorAction SilentlyContinue
        if($check){ return $true }
    }
    return $false
}

# =========================================
# ✅ FINAL ALIAS LOGIC (UPDATED 🔥)
# =========================================
function Get-UniqueAlias {
    param($FirstName,$LastName)

    $maxEmailLength = 30
    $domain = "@coforge.com"
    $maxAliasLength = $maxEmailLength - $domain.Length

    $FirstName = ($FirstName -replace '\s','').ToLower()
    $LastName  = ($LastName -replace '\s','').ToLower()

    # ✅ STEP 1 — Full name
    if ($LastName) {
        $fullAlias = "$FirstName.$LastName"

        if ($fullAlias.Length -le $maxAliasLength -and -not (Alias-Exists $fullAlias)) {
            return $fullAlias
        }
    }

    # ✅ STEP 2 — Short format
    if ($LastName) {
        $baseAlias = "$FirstName.$($LastName.Substring(0,1))"
    } else {
        $baseAlias = $FirstName
    }

    if ($baseAlias.Length -gt $maxAliasLength) {
        $baseAlias = $FirstName.Substring(0,$maxAliasLength)
    }

    if (-not (Alias-Exists $baseAlias)) {
        return $baseAlias
    }

    # ✅ STEP 3 — Numbering (SAFE)
    $count = 1

    while ($true) {

        $newAlias = "$FirstName.$count.$($LastName.Substring(0,1))"

        if ($newAlias.Length -le $maxAliasLength) {

            if (-not (Alias-Exists $newAlias)) {
                return $newAlias
            }
        }

        $count++
    }
}

# =========================================
# CREATE USER
# =========================================
$Password = Generate-RandomPassword
$SecurePassword = ConvertTo-SecureString $Password -AsPlainText -Force

$Alias = Get-UniqueAlias $FirstName $LastName
$UPN = "$Alias@coforge.com"
$Routing = "$Alias@ntlgnoida.mail.onmicrosoft.com"

try {

    Write-Host "Creating user: $DisplayName ($Alias)" -ForegroundColor Cyan

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

    # ✅ ATTRIBUTES
    $CustomAttr1 = "$EmpCode,P"

    Set-RemoteMailbox `
        -Identity $UPN `
        -CustomAttribute1 $CustomAttr1 `
        -CustomAttribute4 $License `
        -EmailAddressPolicyEnabled $false `
        -ErrorAction Stop

    Write-Host "✅ User Created Successfully" -ForegroundColor Green

    # ✅ LOG
    $LogPath = "$LogFolder\$EmpCode.csv"

    [PSCustomObject]@{
        EmpCode=$EmpCode
        DisplayName=$DisplayName
        Alias=$Alias
        Email=$UPN
        Password=$Password
        License=$License
        HRName=$HRName
        Status="SUCCESS"
    } | Export-Csv $LogPath -NoTypeInformation -Encoding UTF8
}
catch {

    $ErrorLog = "$LogFolder\$EmpCode`_ERROR.csv"

    [PSCustomObject]@{
        EmpCode=$EmpCode
        DisplayName=$DisplayName
        Status="FAILED"
        Error=$_.Exception.Message
    } | Export-Csv $ErrorLog -NoTypeInformation -Encoding UTF8
}

if ($Session){ Remove-PSSession $Session }

Get-Job | Stop-Job -ErrorAction SilentlyContinue
Get-Job | Remove-Job -ErrorAction SilentlyContinue

exit