# vybe-test: powershell/host_api/get_host_ui
if (-not $Host.UI) {
    Write-Host "FAIL: expected host UI"
    exit 1
}
Write-Host 'PASS'
exit 0
