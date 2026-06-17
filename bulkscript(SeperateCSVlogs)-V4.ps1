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

# ✅ TIMESTAMP (ONE FILE PER RUN 🔥)
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

    $domains=@("IN.COFORGETECH.COM","UK.COFORGETECH.COM","US.COFORGETECH.COM")

    foreach($domain in $domains){
        $check = Get-ADUser -Filter "SamAccountName -eq '$Alias'" -Server $domain -ErrorAction SilentlyContinue
        if($check){ return $true }
    }
    return $false
}

# =========================================
# ✅ FINAL ALIAS FUNCTION (FIXED 🔥)
# =========================================
function Get-UniqueAlias {
    param($FirstName, $LastName)

    $maxEmailLength = 30
    $domain = "@coforge.com"
    $maxAliasLength = $maxEmailLength - $domain.Length

    # ✅ CLEAN INPUT
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

    # =========================================
    # ✅ STEP 1 — FULL NAME
    # =========================================
    if ($LastName) {

        $baseFull = "$FirstName.$LastName"

        if (Is-AliasAvailable $baseFull) {
            return $baseFull
        }

        # ✅ STEP 2 — FULL NAME NUMBERING
        $count = 1
        while ($true) {

            $newFull = "$FirstName.$count.$LastName"

            # ❌ STOP if exceeding length
            if ($newFull.Length -gt $maxAliasLength) {
                break
            }

            if (Is-AliasAvailable $newFull) {
                return $newFull
            }

            $count++
        }
    }

    # =========================================
    # ✅ STEP 3 — SHORT FORMAT (firstname.l)
    # =========================================
    if ($LastName) {
        $lastInitial = $LastName.Substring(0,1)
        $baseShort = "$FirstName.$lastInitial"
    } else {
        $baseShort = $FirstName
    }

    # ✅ Trim if still long
    if ($baseShort.Length -gt $maxAliasLength) {
        $baseShort = $FirstName.Substring(0, $maxAliasLength)
    }

    if (Is-AliasAvailable $baseShort) {
        return $baseShort
    }

    # =========================================
    # ✅ STEP 4 — SHORT NUMBERING
    # =========================================
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

    # =========================================
    # ✅ FINAL SAFETY FALLBACK (RARE CASE)
    # =========================================
    return $FirstName.Substring(0, [Math]::Min($FirstName.Length, $maxAliasLength))
}

# =========================================
# MAIN LOOP
# =========================================
foreach ($user in $Users) {

    try {

        $FirstName = $user.FirstName
        $LastName = $user.LastName
        $DisplayName = $user.DisplayName
        $EmpCode = $user.EmpCode
        $OUName = $user.OU
        $License = $user.License

        $OUPath = $OUMap.$OUName
        if (!$OUPath) { throw "Invalid OU: $OUName" }

        $Password = Generate-RandomPassword
        $SecurePassword = ConvertTo-SecureString $Password -AsPlainText -Force

        $Alias = Get-UniqueAlias $FirstName $LastName
        $UPN = "$Alias@coforge.com"
        $Routing = "$Alias@ntlgnoida.mail.onmicrosoft.com"

        # ✅ SCREEN LOG
        Write-Host ""
        Write-Host "--------------------------------------------" -ForegroundColor DarkGray
        Write-Host "STARTING USER CREATION" -ForegroundColor Cyan
        Write-Host "Name     : $DisplayName"
        Write-Host "EmpCode  : $EmpCode"
        Write-Host "Alias    : $Alias"
        Write-Host "License  : $License"
        Write-Host "OU       : $OUName"

        # =================================
        # CREATE MAILBOX
        # =================================
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

        Write-Host "✅ Mailbox Created" -ForegroundColor Green

        Start-Sleep 15

        # =================================
        # SET ATTRIBUTES
        # =================================
        $CustomAttr1 = "$EmpCode,P"

        Set-RemoteMailbox `
            -Identity $UPN `
            -CustomAttribute1 $CustomAttr1 `
            -CustomAttribute4 $License `
            -EmailAddressPolicyEnabled $false `
            -ErrorAction Stop

        Write-Host "✅ Attributes Applied" -ForegroundColor Green

        # =================================
        # SUCCESS LOG ✅ (FIXED)
        # =================================
        [PSCustomObject]@{
            DisplayName = $DisplayName
            Alias = $Alias
            Email = $UPN
            EmpCode = $EmpCode
            Password = $Password
            License = $License
            Attribute1 = $CustomAttr1
            Attribute4 = $License
            Status = "SUCCESS"
        } | Export-Csv $SuccessFile -Append -NoTypeInformation

    }
    catch {

        Write-Host "❌ ERROR: $DisplayName" -ForegroundColor Red

        [PSCustomObject]@{
            DisplayName = $DisplayName
            EmpCode = $EmpCode
            Error = $_.Exception.Message
            Status = "FAILED"
        } | Export-Csv $ErrorFile -Append -NoTypeInformation
    }
}

Write-Host ""
Write-Host "============================================="
Write-Host "BULK PROVISIONING COMPLETED ✅"
Write-Host "============================================="

# ✅ CLEANUP
if ($Session){ Remove-PSSession $Session }

exit
