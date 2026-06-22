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
# ✅ EXCHANGE CONNECT (ALWAYS ✅)
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
# ALIAS CHECK ✅ (UNCHANGED)
# =========================================
function Alias-Exists {
    param($Alias)

    $UPN = "$Alias@coforge.com"

    Write-Host ""
    Write-Host "🔍 Checking Alias: $Alias" -ForegroundColor Cyan

    $domains = @("IN.COFORGETECH.COM","UK.COFORGETECH.COM","US.COFORGETECH.COM")

    foreach ($domain in $domains) {
        Write-Host "   ➤ Checking AD ($domain)..."
        if (Get-ADUser -Filter "SamAccountName -eq '$Alias'" -Server $domain -ErrorAction SilentlyContinue) {
            Write-Host "   ⚠ FOUND in AD: $Alias ($domain)" -ForegroundColor Yellow
            return $true
        }
    }

    Write-Host "   ➤ Checking Exchange Mailbox..."
    if (Get-Mailbox -Identity $UPN -ErrorAction SilentlyContinue) {
        Write-Host "   ⚠ FOUND in Exchange Mailbox: $UPN" -ForegroundColor Yellow
        return $true
    }

    Write-Host "   ➤ Checking RemoteMailbox..."
    if (Get-RemoteMailbox -Identity $UPN -ErrorAction SilentlyContinue) {
        Write-Host "   ⚠ FOUND in RemoteMailbox: $UPN" -ForegroundColor Yellow
        return $true
    }

    Write-Host "   ➤ Checking Exchange Alias..."
    if (Get-Recipient -Filter "Alias -eq '$Alias'" -ErrorAction SilentlyContinue) {
        Write-Host "   ⚠ FOUND in Exchange Alias: $Alias" -ForegroundColor Yellow
        return $true
    }

    Write-Host "   ➤ Checking Azure AD..."
    try {
        if (Get-MgUser -Filter "userPrincipalName eq '$UPN'" -ErrorAction SilentlyContinue) {
            Write-Host "   ⚠ FOUND in Azure AD: $UPN" -ForegroundColor Yellow
            return $true
        }
    } catch {}

    Write-Host "   ✅ NOT FOUND anywhere: $Alias" -ForegroundColor Green
    return $false
}

# =========================================
# ALIAS GENERATION ✅
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

    return $FirstName
}

# =========================================
# MAIN LOOP (UPDATED ✅)
# =========================================
foreach($user in $Users){

    try {

        $FirstName=$user.FirstName
        $LastName=$user.LastName
        $EmpCodeRaw=$user.EmpCode
        $OUName=$user.OU
        $License=$user.License

        # ✅ EMP 8 DIGIT FORMAT
        $EmpCode = $EmpCodeRaw.ToString().PadLeft(8,'0')
        $CustomAttr1 = "$EmpCode,P"

        Write-Host ""
        Write-Host "🔍 Checking EmpCode in AD: $CustomAttr1" -ForegroundColor Cyan

        $EmpCheck = Get-ADUser -Filter "Description -eq '$CustomAttr1'" -ErrorAction SilentlyContinue

        if ($EmpCheck) {
            Write-Host "❌ SKIPPED: EmpCode already exists → $CustomAttr1" -ForegroundColor Red

            [PSCustomObject]@{
                DisplayName = "$FirstName $LastName"
                Alias       = "N/A"
                Email       = "N/A"
                EmpCode     = $EmpCode
                Password    = "N/A"
                License     = $License
                Attribute1  = $CustomAttr1
                Status      = "EMP_DUPLICATE"
            } | Export-Csv $ErrorFile -Append -NoTypeInformation

            continue
        }

        $OUPath=$OUMap.$OUName
        if(!$OUPath){ throw "Invalid OU: $OUName" }

        $DisplayName="$FirstName $LastName"

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

        # ✅ LOG
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
    }
}

Write-Host ""
Write-Host "✅ BULK PROCESS COMPLETED ✅"

# ✅ FIXED RERUN
& $MyInvocation.MyCommand.Path
