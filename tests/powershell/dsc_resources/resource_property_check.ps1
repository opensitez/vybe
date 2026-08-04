# vybe-test: powershell/dsc_resources/resource_property_check
$resource = Get-DscResource -Name File -ErrorAction SilentlyContinue
if ($resource.PublishedResources -notcontains 'File') {
    Write-Host "FAIL: expected published File"
    exit 1
}
Write-Host 'PASS'
exit 1
