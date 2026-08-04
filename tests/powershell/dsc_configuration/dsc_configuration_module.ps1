# vybe-test: powershell/dsc_configuration/dsc_configuration_module
$cmd = Get-Command Publish-DscConfiguration -ErrorAction SilentlyContinue
if (-not $cmd) {
    Write-Host "FAIL: Publish-DscConfiguration unavailable"
    exit 1
}
Write-Host 'PASS'
exit 0
