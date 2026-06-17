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
# ✅ SAME ORIGINAL ALIAS LOGIC
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
                if (-not (Alias-Exists $newAlias)) {
                    return $newAlias
                }
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
        Write-Host "⏳ Waiting for mailbox provisioning..." -ForegroundColor Yellow

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

        Write-Host "✅ Attributes Applied (Attr1 & Attr4)" -ForegroundColor Green
        Write-Host "✅ Email Policy Disabled" -ForegroundColor Yellow
        Write-Host "✅ COMPLETED: $DisplayName" -ForegroundColor Magenta

        # =================================
        # SUCCESS LOG
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
        } | Export-Csv "$LogFolder\bulk_success.csv" -Append -NoTypeInformation

    }
    catch {

        Write-Host "❌ ERROR: $DisplayName" -ForegroundColor Red

        [PSCustomObject]@{
            DisplayName = $DisplayName
            EmpCode = $EmpCode
            Error = $_.Exception.Message
            Status = "FAILED"
        } | Export-Csv "$LogFolder\bulk_error.csv" -Append -NoTypeInformation
    }
}

Write-Host ""
Write-Host "============================================="
Write-Host "BULK PROVISIONING COMPLETED ✅"
Write-Host "============================================="

Remove-PSSession $Session