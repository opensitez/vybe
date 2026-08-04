# vybe-test: powershell/host_api/get_host_name
if (-not $Host.Name) {
    Write-Host "FAIL: expected host name"
    exit 1
}
Write-Host 'PASS'
exit 0
