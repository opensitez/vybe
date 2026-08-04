# vybe-test: powershell/dsc_configuration/test_dsc_configuration
if (-not (Get-Command Test-DscConfiguration -ErrorAction SilentlyContinue)) {
    Write-Host "FAIL: Test-DscConfiguration unavailable"
    exit 1
}
Write-Host 'PASS'
exit 0
