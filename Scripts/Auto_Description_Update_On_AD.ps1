Import-Module ActiveDirectory

# =========================================
# LOG FOLDER
# =========================================
$LogFolder = "C:\Users\Gaurav.26\OneDrive - Coforge Limited\Documents\UserAutomation\Logs"

# ✅ GET LATEST CSV
$csvFile = Get-ChildItem -Path $LogFolder -Filter "bulk_success_*.csv" |
           Sort-Object LastWriteTime -Descending |
           Select-Object -First 1

if (-not $csvFile) {
    Write-Host "❌ No CSV file found" -ForegroundColor Red
    Read-Host "Press Enter to exit..."
    exit
}

Write-Host "Using File: $($csvFile.Name)" -ForegroundColor Yellow

$users = Import-Csv $csvFile.FullName

# ✅ ADMIN CREDENTIAL (VERY IMPORTANT)
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

        $updated = $false

        foreach ($domain in $domains) {

            try {

                # ✅ Find user in correct domain
                $adUser = Get-ADUser `
                    -Identity $alias `
                    -Server $domain `
                    -Credential $cred `
                    -ErrorAction Stop

                Write-Host "[FOUND] $alias in $domain" -ForegroundColor Cyan

                # ✅ Update description
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

                $updated = $true
                break
            }
            catch {
                Write-Host "[NOT FOUND] $alias in $domain" -ForegroundColor DarkYellow
            }
        }

        if (-not $updated) {
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
