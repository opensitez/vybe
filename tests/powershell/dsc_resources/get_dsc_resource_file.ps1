# vybe-test: powershell/dsc_resources/get_dsc_resource_file
$cmd = Get-Command Get-DscResource -ErrorAction SilentlyContinue
if (-not $cmd) {
    Write-Host "FAIL: Get-DscResource unavailable"
    exit 1
}
$resource = Get-DscResource -Name File -ErrorAction SilentlyContinue
if (-not $resource) {
    Write-Host "FAIL: expected File resource"
    exit 1
}
Write-Host 'PASS'
exit 0
