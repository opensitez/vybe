# vybe-test: powershell/host_api/host_api_is_readonly
if ($Host.IsRunspacePushed -ne $false -and $Host.IsRunspacePushed -ne $true) {
    Write-Host "FAIL: expected boolean runspace pushed property"
    exit 1
}
Write-Host 'PASS'
exit 0
