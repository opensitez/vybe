# vybe-test: powershell/dsc_configuration/publish_dsc_configuration
if (-not (Get-Command Publish-DscConfiguration -ErrorAction SilentlyContinue)) {
    Write-Host "FAIL: Publish-DscConfiguration unavailable"
    exit 1
}
Write-Host 'PASS'
exit 0
