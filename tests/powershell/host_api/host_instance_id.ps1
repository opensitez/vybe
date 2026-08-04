# vybe-test: powershell/host_api/host_instance_id
if (-not $Host.InstanceId) {
    Write-Host "FAIL: expected instance id"
    exit 1
}
Write-Host 'PASS'
exit 0
