# =========================================
# BASE FOLDER
# =========================================

$BaseFolder = "C:\Users\Gaurav.26\OneDrive - Coforge Limited\Documents\UserAutomation"

# =========================================
# FOLDER STRUCTURE
# =========================================

$queueFolder     = "$BaseFolder\Input_JSON"
$processedFolder = "$BaseFolder\Processed_JSON"
$logFolder       = "$BaseFolder\Logs"

$createUserScript = "$BaseFolder\Scripts\CreateUserV2.ps1"

Write-Host ""
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "HR AUTOMATION WATCHER STARTED" -ForegroundColor Green
Write-Host "=============================================" -ForegroundColor Cyan

# =========================================
# CHECK FOLDERS
# =========================================

foreach ($folder in @($queueFolder, $processedFolder, $logFolder)) {
    if (!(Test-Path $folder)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }
}

# =========================================
# GET JSON FILES
# =========================================

$jsonFiles = Get-ChildItem -Path $queueFolder -Filter *.json -ErrorAction SilentlyContinue

if (!$jsonFiles -or $jsonFiles.Count -eq 0) {
    Write-Host ""
    Write-Host "NO JSON FILE FOUND" -ForegroundColor Yellow
    Start-Sleep 5
    exit
}

# =========================================
# PROCESS FILES
# =========================================

foreach ($jsonFile in $jsonFiles) {

    try {

        Write-Host ""
        Write-Host "PROCESSING FILE : $($jsonFile.Name)" -ForegroundColor Cyan

        $jsonData = Get-Content $jsonFile.FullName -Raw | ConvertFrom-Json

        # ✅ READ JSON
        $FirstName   = $jsonData.FirstName
        $LastName    = $jsonData.LastName
        $DisplayName = $jsonData.DisplayName
        $OU          = $jsonData.OU
        $EmpCode     = $jsonData.EmployeeCode
        $License     = $jsonData.License
        $HRName      = $jsonData.HR

        Write-Host "EmpCode : $EmpCode"
        Write-Host "License : $License" -ForegroundColor Yellow   # ✅ Debug

        # =====================================
        # RUN CREATE USER ✅ FIXED
        # =====================================

        Write-Host "Starting CreateUser Script..." -ForegroundColor Yellow

        powershell.exe -ExecutionPolicy Bypass -File $createUserScript `
            -Mode "Single" `
            -FirstName "$FirstName" `
            -LastName "$LastName" `
            -DisplayName "$DisplayName" `
            -OU "$OU" `
            -EmpCode "$EmpCode" `
            -License "$License"

        Write-Host "CreateUser execution completed ✅" -ForegroundColor Green

        # =====================================
        # WAIT FOR CSV ✅
        # =====================================

        $SafeName = $EmpCode
        $SuccessLog = "$logFolder\$SafeName.csv"
        $ErrorLog   = "$logFolder\$SafeName`_ERROR.csv"

        $maxWait = 60
        $elapsed = 0

        while ($elapsed -lt $maxWait) {

            if ((Test-Path $SuccessLog) -or (Test-Path $ErrorLog)) {

                Write-Host ""
                Write-Host "CSV LOG DETECTED ✅" -ForegroundColor Green
                break
            }

            Start-Sleep 2
            $elapsed += 2
        }

    }
    catch {

        Write-Host ""
        Write-Host "WATCHER ERROR ❌" -ForegroundColor Red
        Write-Host $_.Exception.Message
    }

    finally {

        # =====================================
        # ALWAYS MOVE JSON ✅ (IMPORTANT FIX)
        # =====================================

        try {
            if (Test-Path $jsonFile.FullName) {

                $destination = "$processedFolder\$($jsonFile.Name)"

                Move-Item `
                    -Path $jsonFile.FullName `
                    -Destination $destination `
                    -Force

                Write-Host ""
                Write-Host "JSON MOVED TO PROCESSED ✅" -ForegroundColor Green
            }
        }
        catch {
            Write-Host "Failed to move JSON ❌" -ForegroundColor Red
        }
    }
}

Write-Host ""
Write-Host "============================================="
Write-Host "PROCESS COMPLETED ✅"
Write-Host "============================================="

Start-Sleep 5

Get-Job | Stop-Job -ErrorAction SilentlyContinue
Get-Job | Remove-Job -ErrorAction SilentlyContinue

exit