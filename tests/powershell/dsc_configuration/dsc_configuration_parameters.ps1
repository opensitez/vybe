# vybe-test: powershell/dsc_configuration/dsc_configuration_parameters
$cmd = Get-Command Test-DscConfiguration -ErrorAction SilentlyContinue
if (-not $cmd) {
    Write-Host "FAIL: Test-DscConfiguration unavailable"
    exit 1
}
Write-Host 'PASS'
exit 0
