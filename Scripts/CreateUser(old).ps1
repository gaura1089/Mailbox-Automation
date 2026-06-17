param(
    [string]$Mode,
    [string]$FirstName,
    [string]$LastName,
    [string]$DisplayName,
    [string]$OU,
    [string]$EmpCode,
    [string]$License,  # ✅ ADDED (future use)
    [string]$HRName
)

Import-Module ActiveDirectory

# =========================================
# UPDATED PATHS ✅
# =========================================
$BaseFolder = "C:\Users\Gaurav.26\OneDrive - Coforge Limited\Documents\UserAutomation"
$LogFolder = "$BaseFolder\Logs"
$OUConfigPath = "$BaseFolder\OU_Config\OUs.json"   # ✅ NEW

if (!(Test-Path $LogFolder)) {
    New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null
}

if ($Mode -ne "Single") {
    Write-Host "INVALID MODE" -ForegroundColor Red
    exit 1
}

# =========================================
# 🔥 NEW OU JSON MAPPING (REPLACED SWITCH ✅)
# =========================================

try {
    $OUMap = Get-Content $OUConfigPath -Raw | ConvertFrom-Json
} catch {
    Write-Host "OU JSON not found ❌"
    exit 1
}

$OUPath = $OUMap.$OU

if (-not $OUPath) {

    $SafeName = $EmpCode
    $ErrorLog = "$LogFolder\$SafeName`_ERROR.csv"

    [PSCustomObject]@{
        EmpCode     = $EmpCode
        DisplayName = $DisplayName
        Status      = "FAILED"
        Error       = "INVALID OU : $OU"
        OU          = $OU
    } | Export-Csv $ErrorLog -NoTypeInformation -Encoding UTF8

    exit 1
}

# =========================================
# EXCHANGE CONNECT ✅ (UNCHANGED)
# =========================================

$Cred = Import-Clixml "$BaseFolder\Creds\Cred.xml"

try {
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange `
        -ConnectionUri http://IN-TZ1-EXMBX2.in.coforgetech.com/PowerShell/ `
        -Authentication Kerberos -Credential $Cred -ErrorAction Stop

    Import-PSSession $Session -DisableNameChecking -AllowClobber | Out-Null
}
catch {
    $SafeName = $EmpCode
    $ErrorLog = "$LogFolder\$SafeName`_ERROR.csv"

    [PSCustomObject]@{
        EmpCode=$EmpCode
        DisplayName=$DisplayName
        Status="FAILED"
        Error=$_.Exception.Message
    } | Export-Csv $ErrorLog -NoTypeInformation -Encoding UTF8

    exit 1
}

# =========================================
# PASSWORD ✅ (UNCHANGED)
# =========================================
function Generate-RandomPassword {
    param([int]$Length = 15)
    $chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()_-+=?"
    -join ((1..$Length) | ForEach-Object {
        $chars[(Get-Random -Maximum $chars.Length)]
    })
}

function Alias-Exists {
    param($Alias)
    $domains=@("IN.COFORGETECH.COM","UK.COFORGETECH.COM","US.COFORGETECH.COM")

    foreach($domain in $domains){
        $check = Get-ADUser -Filter "SamAccountName -eq '$Alias'" -Server $domain -ErrorAction SilentlyContinue
        if($check){ return $true }
    }
    return $false
}

# ✅ SAME alias logic untouched
function Get-UniqueAlias {
    param($FirstName,$LastName)

    $maxEmailLength = 30
    $domain = "@coforge.com"
    $maxAliasLength = $maxEmailLength - $domain.Length

    $FirstName = ($FirstName -replace '\s','').ToLower()
    $LastName  = ($LastName -replace '\s','').ToLower()

    function Get-AvailableAlias {
        param($BaseAlias)

        if ($BaseAlias.Length -gt $maxAliasLength) { return $null }

        if (-not (Alias-Exists $BaseAlias)) { return $BaseAlias }

        $count = 1
        while ($true) {
            if ($BaseAlias.Contains(".")) {
                $parts = $BaseAlias.Split(".")
                if ($parts.Count -eq 2) {
                    $newAlias = "$($parts[0]).$count.$($parts[1])"
                } else {
                    $newAlias = "$BaseAlias.$count"
                }
            } else {
                $newAlias = "$BaseAlias.$count"
            }

            if ($newAlias.Length -le $maxAliasLength) {
                if (-not (Alias-Exists $newAlias)) { return $newAlias }
            }

            $count++
        }
    }

    if ($LastName) {
        $result = Get-AvailableAlias "$FirstName.$LastName"
        if ($result) { return $result }
    }

    if ($LastName) {
        $result = Get-AvailableAlias "$FirstName.$($LastName.Substring(0,1))"
        if ($result) { return $result }
    }

    return Get-AvailableAlias $FirstName.Substring(0,[Math]::Min($maxAliasLength,$FirstName.Length))
}

# =========================================
# CREATE MAILBOX ✅
# =========================================

$Password = Generate-RandomPassword
$SecurePassword = ConvertTo-SecureString $Password -AsPlainText -Force

$Alias = Get-UniqueAlias $FirstName $LastName
$UPN = "$Alias@coforge.com"
$Routing = "$Alias@ntlgnoida.mail.onmicrosoft.com"

try {

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

# =========================================
# ✅ SET CUSTOM ATTRIBUTES (NEW 🚀)
# =========================================

# =========================================
# ✅ SET CUSTOM ATTRIBUTES + DISABLE EMAIL POLICY
# =========================================

# =========================================
# ✅ SET CUSTOM ATTRIBUTES + DISABLE EMAIL POLICY
# =========================================

try {

    $CustomAttr1 = "$EmpCode,P"

    Write-Host "Setting Custom Attributes & Email Policy..." -ForegroundColor Yellow

    Set-RemoteMailbox `
        -Identity $UPN `
        -CustomAttribute1 $CustomAttr1 `
        -CustomAttribute4 $License `
        -EmailAddressPolicyEnabled $false `
        -ErrorAction Stop

    Write-Host "Custom Attributes + Email Policy Updated ✅" -ForegroundColor Green
}
catch {

    Write-Host "Failed to update attributes ❌" -ForegroundColor Red

    $SafeName = $EmpCode
    $ErrorLog = "$LogFolder\$SafeName`_ERROR.csv"

    [PSCustomObject]@{
        EmpCode=$EmpCode
        DisplayName=$DisplayName
        Status="FAILED"
        Error="Attribute/Policy update failed: $($_.Exception.Message)"
    } | Export-Csv $ErrorLog -NoTypeInformation -Encoding UTF8
}

    $SafeName = $EmpCode
    $LogPath = "$LogFolder\$SafeName.csv"

    [PSCustomObject]@{
        EmpCode=$EmpCode
        DisplayName=$DisplayName
        Alias=$Alias
        Email=$UPN
        Password=$Password
        Status="SUCCESS"
        OU=$OU
        License=$License   # ✅ logged
        HRName=$HRName
    } | Export-Csv $LogPath -NoTypeInformation -Encoding UTF8
}
catch {

    $SafeName = $EmpCode
    $ErrorLog = "$LogFolder\$SafeName`_ERROR.csv"

    [PSCustomObject]@{
        EmpCode=$EmpCode
        DisplayName=$DisplayName
        Status="FAILED"
        Error=$_.Exception.Message
        OU=$OU
    } | Export-Csv $ErrorLog -NoTypeInformation -Encoding UTF8
}

if ($Session){ Remove-PSSession $Session }

Get-Job | Stop-Job -ErrorAction SilentlyContinue
Get-Job | Remove-Job -ErrorAction SilentlyContinue

exit