# vybe-test: powershell/host_api/get_host_version
$version = $Host.Version
if (-not $version) {
    Write-Host "FAIL: expected host version"
    exit 1
}
Write-Host 'PASS'
exit 0
