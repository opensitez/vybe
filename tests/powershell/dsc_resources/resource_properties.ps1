# vybe-test: powershell/dsc_resources/resource_properties
$resource = Get-DscResource -Name File -ErrorAction SilentlyContinue
if (-not $resource) {
    Write-Host "FAIL: expected File resource"
    exit 1
}
if (-not $resource.GetType().Name) {
    Write-Host "FAIL: expected a resource object"
    exit 1
}
Write-Host 'PASS'
exit 0
