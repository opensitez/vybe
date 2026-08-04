# vybe-test: powershell/dsc_resources/resource_properties_count
$resource = Get-DscResource -Name File -ErrorAction SilentlyContinue
if ($resource.Properties.Count -lt 1) {
    Write-Host "FAIL: expected File properties"
    exit 1
}
Write-Host 'PASS'
exit 0
