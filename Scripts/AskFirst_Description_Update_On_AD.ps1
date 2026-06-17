Import-Module ActiveDirectory

# =========================================
# ASK USER FOR CSV FILE
# =========================================
$csvPath = Read-Host "Enter full CSV file path"

if (!(Test-Path $csvPath)) {
    Write-Host "❌ File not found. Please check path." -ForegroundColor Red
    Read-Host "Press Enter to exit..."
    exit
}

Write-Host "Using File: $csvPath" -ForegroundColor Yellow

# =========================================
# IMPORT CSV
# =========================================
$users = Import-Csv $csvPath

# ✅ ADMIN CREDENTIAL
$cred = Get-Credential -Message "Enter AD Admin Credentials"

# ✅ DOMAINS
$domains = @(
    "IN.COFORGETECH.COM",
    "US.COFORGETECH.COM",
    "UK.COFORGETECH.COM"
)

foreach ($user in $users) {

    try {

        $alias = $user.Alias
        $attr1 = $user.Attribute1

        Write-Host ""
        Write-Host "Processing: $alias" -ForegroundColor Cyan

        if (-not $attr1) {
            Write-Host "⚠ Attribute1 empty — skipping" -ForegroundColor Yellow
            continue
        }

        $found = $false

        foreach ($domain in $domains) {

            try {

                # ✅ Find user
                $adUser = Get-ADUser `
                    -Identity $alias `
                    -Server $domain `
                    -Credential $cred `
                    -ErrorAction Stop

                Write-Host "[FOUND] $alias in $domain" -ForegroundColor Cyan

                # ✅ Update Description
                Set-ADUser `
                    -Identity $alias `
                    -Description $attr1 `
                    -Server $domain `
                    -Credential $cred `
                    -ErrorAction Stop

                Start-Sleep 2

                # ✅ VERIFY
                $verifyUser = Get-ADUser `
                    -Identity $alias `
                    -Server $domain `
                    -Credential $cred `
                    -Properties Description

                if ($verifyUser.Description -eq $attr1) {
                    Write-Host "✅ VERIFIED ✅ Description updated: $($verifyUser.Description)" -ForegroundColor Green
                }
                else {
                    Write-Host "❌ Verification failed" -ForegroundColor Red
                }

                $found = $true
                break
            }
            catch {
                Write-Host "[NOT FOUND] $alias in $domain" -ForegroundColor DarkYellow
            }
        }

        if (-not $found) {
            Write-Host "❌ User not found in any domain: $alias" -ForegroundColor Red
        }

    }
    catch {
        Write-Host "❌ ERROR for: $alias" -ForegroundColor Red
        Write-Host $_.Exception.Message
    }
}

Write-Host ""
Write-Host "=========================================="
Write-Host "DESCRIPTION UPDATE COMPLETED ✅"
Write-Host "=========================================="

Read-Host "Press Enter to exit..."
exit