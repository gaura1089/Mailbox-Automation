Import-Module ActiveDirectory

# =========================================
# PATHS
# =========================================
$BaseFolder = "C:\Users\Gaurav.26\OneDrive - Coforge Limited\Documents\UserAutomation"
$OUConfigPath = "$BaseFolder\OU_Config\OUs.json"
$LogFolder = "$BaseFolder\Logs"

if (!(Test-Path $LogFolder)) {
    New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null
}

# =========================================
# CSV INPUT
# =========================================
$csvPath = "$BaseFolder\bulk_users.csv"
$Users = Import-Csv $csvPath

# =========================================
# LOAD OU MAP
# =========================================
$OUMap = Get-Content $OUConfigPath -Raw | ConvertFrom-Json

# =========================================
# EXCHANGE SESSION
# =========================================
$Cred = Import-Clixml "$BaseFolder\Creds\Cred.xml"

$Session = New-PSSession -ConfigurationName Microsoft.Exchange `
    -ConnectionUri http://IN-TZ1-EXMBX2.in.coforgetech.com/PowerShell/ `
    -Authentication Kerberos -Credential $Cred

Import-PSSession $Session -DisableNameChecking -AllowClobber | Out-Null

# =========================================
# PASSWORD FUNCTION
# =========================================
function Generate-RandomPassword {
    param([int]$Length = 12)
    $chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%"
    -join ((1..$Length) | ForEach-Object {
        $chars[(Get-Random -Maximum $chars.Length)]
    })
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

        if (!$OUPath) {
            throw "Invalid OU: $OUName"
        }

        # PASSWORD
        $Password = Generate-RandomPassword
        $SecurePassword = ConvertTo-SecureString $Password -AsPlainText -Force

        # ALIAS
        $Alias = "$($FirstName.ToLower()).$($LastName.ToLower())"
        $UPN = "$Alias@coforge.com"
        $Routing = "$Alias@ntlgnoida.mail.onmicrosoft.com"

        Write-Host "Creating user: $DisplayName"

        # CREATE USER
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

        Start-Sleep 20

        # =========================================
        # SET ATTRIBUTES + EMAIL POLICY
        # =========================================

        $CustomAttr1 = "$EmpCode,P"

        Set-RemoteMailbox `
            -Identity $UPN `
            -CustomAttribute1 $CustomAttr1 `
            -CustomAttribute4 $License `
            -EmailAddressPolicyEnabled $false `
            -ErrorAction Stop

        # =========================================
        # LOG SUCCESS
        # =========================================

        $LogPath = "$LogFolder\bulk_success.csv"

        [PSCustomObject]@{
            DisplayName = $DisplayName
            Email = $UPN
            EmpCode = $EmpCode
            Password = $Password
            License = $License
            Attribute1 = $CustomAttr1
            Attribute4 = $License
            Status = "SUCCESS"
        } | Export-Csv $LogPath -Append -NoTypeInformation

    }
    catch {

        $ErrorPath = "$LogFolder\bulk_error.csv"

        [PSCustomObject]@{
            DisplayName = $DisplayName
            EmpCode = $EmpCode
            Error = $_.Exception.Message
            Status = "FAILED"
        } | Export-Csv $ErrorPath -Append -NoTypeInformation

    }
}

Write-Host "Bulk provisioning completed ✅"
