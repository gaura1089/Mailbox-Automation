Import-Module ActiveDirectory

# =========================================
# PATHS
# =========================================
$BaseFolder = "C:\Users\Gaurav.26\OneDrive - Coforge Limited\Documents\UserAutomation"
$OUConfigPath = "$BaseFolder\OU_Config\OUs.json"
$LogFolder = "$BaseFolder\Logs"
$csvPath = "$BaseFolder\bulk_users.csv"

if (!(Test-Path $LogFolder)) {
    New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null
}

$Users = Import-Csv $csvPath
$OUMap = Get-Content $OUConfigPath -Raw | ConvertFrom-Json

# ✅ TIMESTAMP
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$SuccessFile = "$LogFolder\bulk_success_$timestamp.csv"
$ErrorFile   = "$LogFolder\bulk_error_$timestamp.csv"

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
    param([int]$Length = 12)
    $chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%"
    -join ((1..$Length) | ForEach-Object {
        $chars[(Get-Random -Maximum $chars.Length)]
    })
}

# =========================================
# DOMAIN CHECK
# =========================================
function Alias-Exists {
    param($Alias)

    $domains = @("IN.COFORGETECH.COM","UK.COFORGETECH.COM","US.COFORGETECH.COM")

    # ✅ AD CHECK
    foreach ($domain in $domains) {
        $adUser = Get-ADUser -Filter "SamAccountName -eq '$Alias'" -Server $domain -ErrorAction SilentlyContinue
        if ($adUser) { return $true }
    }

    # ✅ EXCHANGE CHECK (IMPORTANT 🔥)
    $recipient = Get-Recipient -Filter "Alias -eq '$Alias'" -ErrorAction SilentlyContinue
    if ($recipient) { return $true }

    return $false
}

# =========================================
# ✅ FINAL ALIAS FUNCTION (HYBRID + SAFE)
# =========================================
function Get-UniqueAlias {
    param($FirstName, $LastName)

    $maxEmailLength = 30
    $domain = "@coforge.com"
    $maxAliasLength = $maxEmailLength - $domain.Length

    $FirstName = ($FirstName -replace '\s','').ToLower()

    if ([string]::IsNullOrWhiteSpace($LastName)) {
        $LastName = ""
    } else {
        $LastName = ($LastName -replace '\s','').ToLower()
    }

    function Is-AliasAvailable {
        param($alias)

        if ($alias.Length -gt $maxAliasLength) {
            return $false
        }
        return -not (Alias-Exists $alias)
    }

    # ✅ FULL NAME
    if ($LastName) {

        $baseFull = "$FirstName.$LastName"

        if (Is-AliasAvailable $baseFull) {
            return $baseFull
        }

        # ✅ FULL NUMBERING
        $count = 1
        while ($true) {

            $newFull = "$FirstName.$count.$LastName"

            if ($newFull.Length -gt $maxAliasLength) {
                break
            }

            if (Is-AliasAvailable $newFull) {
                return $newFull
            }

            $count++
        }
    }

    # ✅ SHORT FORMAT
    if ($LastName) {
        $baseShort = "$FirstName.$($LastName.Substring(0,1))"
    } else {
        $baseShort = $FirstName
    }

    if ($baseShort.Length -gt $maxAliasLength) {
        $baseShort = $FirstName.Substring(0, $maxAliasLength)
    }

    if (Is-AliasAvailable $baseShort) {
        return $baseShort
    }

    # ✅ SHORT NUMBERING
    $count = 1
    while ($true) {

        if ($LastName) {
            $newShort = "$FirstName.$count.$($LastName.Substring(0,1))"
        }
        else {
            $newShort = "$FirstName.$count"
        }

        if ($newShort.Length -le $maxAliasLength) {
            if (Is-AliasAvailable $newShort) {
                return $newShort
            }
        }

        $count++
    }

    return $FirstName.Substring(0, [Math]::Min($FirstName.Length, $maxAliasLength))
}

# =========================================
# MAIN LOOP
# =========================================
foreach ($user in $Users) {

    try {

        $FirstName = $user.FirstName
        $LastName = $user.LastName
        $EmpCode = $user.EmpCode
        $OUName = $user.OU
        $License = $user.License

        # ✅ VALIDATION
        if ([string]::IsNullOrWhiteSpace($FirstName)) {
            throw "FirstName missing in CSV"
        }

        $OUPath = $OUMap.$OUName
        if (!$OUPath) { throw "Invalid OU: $OUName" }

        # ✅ DISPLAY NAME FIX (NO CONFLICT 🔥)
        if ([string]::IsNullOrWhiteSpace($LastName)) {
            $BaseDisplayName = "$FirstName"
        } else {
            $BaseDisplayName = "$FirstName $LastName"
        }

        $DisplayName = $BaseDisplayName

        if (Get-ADUser -Filter "Name -eq '$DisplayName'" -ErrorAction SilentlyContinue) {
            $DisplayName = "$BaseDisplayName - $EmpCode"
        }

        # ✅ PASSWORD
        $Password = Generate-RandomPassword
        $SecurePassword = ConvertTo-SecureString $Password -AsPlainText -Force

        # ✅ ALIAS
        $Alias = Get-UniqueAlias $FirstName $LastName
        $UPN = "$Alias@coforge.com"
        $Routing = "$Alias@ntlgnoida.mail.onmicrosoft.com"

        # ✅ LOG
        Write-Host ""
        Write-Host "--------------------------------------------" -ForegroundColor DarkGray
        Write-Host "Creating: $DisplayName" -ForegroundColor Cyan
        Write-Host "Alias   : $Alias"

        # ✅ CREATE MAILBOX
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

        Start-Sleep 15

        # ✅ ATTRIBUTES
        $CustomAttr1 = "$EmpCode,P"

        Set-RemoteMailbox `
            -Identity $UPN `
            -CustomAttribute1 $CustomAttr1 `
            -CustomAttribute4 $License `
            -EmailAddressPolicyEnabled $false `
            -ErrorAction Stop

        Write-Host "✅ SUCCESS: $DisplayName" -ForegroundColor Green

        # ✅ SUCCESS LOG
        [PSCustomObject]@{
            DisplayName = $DisplayName
            Alias = $Alias
            Email = $UPN
            EmpCode = $EmpCode
            Password = $Password
            License = $License
            Attribute1 = $CustomAttr1
            Status = "SUCCESS"
        } | Export-Csv $SuccessFile -Append -NoTypeInformation
    }
    catch {

        Write-Host "❌ ERROR: $($user.FirstName) $($user.LastName)" -ForegroundColor Red
        Write-Host $_.Exception.Message

        [PSCustomObject]@{
            DisplayName = "$($user.FirstName) $($user.LastName)"
            EmpCode = $user.EmpCode
            Error = $_.Exception.Message
            Status = "FAILED"
        } | Export-Csv $ErrorFile -Append -NoTypeInformation
    }
}

Write-Host ""
Write-Host "============================================="
Write-Host "BULK PROVISIONING COMPLETED ✅"
Write-Host "============================================="

if ($Session){ Remove-PSSession $Session }

exit
