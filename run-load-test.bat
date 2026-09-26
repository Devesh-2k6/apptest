@echo off
title ExpiryGo - Baseline Load Test Runner (100 Users, 1 Minute)
echo ==============================================================================
echo        EXPIRYGO BASELINE LOAD TEST RUNNER (100 VIRTUAL USERS)
echo ==============================================================================
echo.
cd /d "%~dp0\load-tests"

if not exist "node_modules\exceljs" (
    echo [INFO] Installing required dependencies...
    call npm install
)

echo.
echo [1/2] Executing 100 Virtual Users Baseline Load Test (60 Seconds)...
echo.
node run-baseline-load-test.js --users 100 --duration 60

echo.
echo ==============================================================================
echo [2/2] Test Complete! Reports Generated in:
echo       - load-tests\results\ExpiryGo_Baseline_Load_Test_Report.xlsx
echo       - load-tests\results\baseline-load-test-report.md
echo ==============================================================================
pause
