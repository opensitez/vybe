# vybe-test: powershell/host_api/get_host_private_data
if ($Host.PrivateData -eq $null) {
    Write-Host "FAIL: expected private data property"
    exit 1
}
Write-Host 'PASS'
exit 0
