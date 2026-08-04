# vybe-test: powershell/host_api/get_host_version_major
if ($Host.Version.Major -lt 0) {
    Write-Host "FAIL: expected major version"
    exit 1
}
Write-Host 'PASS'
exit 0
