# vybe-test: powershell/automatic_variables/dollar_lastexitcode
$null = pwsh -NoProfile -Command "exit 0"
if ($LASTEXITCODE -ne 0) {
    Write-Host "FAIL: LASTEXITCODE check failed"
    exit 1
}
Write-Host "PASS"
exit 0
