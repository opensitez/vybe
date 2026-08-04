# vybe-test: powershell/dsc_configuration/get_publish_dsc_command
if (-not (Get-Command Publish-DscConfiguration -ErrorAction SilentlyContinue)) {
    Write-Host "FAIL: Publish-DscConfiguration unavailable"
    exit 1
}
Write-Host 'PASS'
exit 0
