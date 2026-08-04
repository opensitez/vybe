# vybe-test: powershell/automatic_variables/dollar_lastexitcode
# Run a simple external command and check $LASTEXITCODE
& powershell -Command "exit 0"
if ($LASTEXITCODE -ne 0) { Write-Host "FAIL: expected 0"; exit 1 }
& powershell -Command "exit 42"
if ($LASTEXITCODE -ne 42) { Write-Host "FAIL: expected 42, got $LASTEXITCODE"; exit 1 }
Write-Host "PASS"
exit 0
