# vybe-test: powershell/dsc_configuration/get_dsc_configuration_command
if (-not (Get-Command Test-DscConfiguration -ErrorAction SilentlyContinue)) {
    Write-Host "FAIL: Test-DscConfiguration unavailable"
    exit 1
}
Write-Host 'PASS'
exit 0
