# vybe-test: powershell/dsc_configuration/dsc_configuration_test_command
$cmd = Get-Command Test-DscConfiguration -ErrorAction SilentlyContinue
if (-not $cmd) {
    Write-Host "FAIL: command unavailable"
    exit 1
}
Write-Host 'PASS'
exit 0
