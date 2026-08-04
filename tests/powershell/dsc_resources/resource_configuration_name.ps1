# vybe-test: powershell/dsc_resources/resource_configuration_name
$resource = Get-DscResource -Name File -ErrorAction SilentlyContinue
if ($resource.Name -ne 'File') {
    Write-Host "FAIL: expected File name"
    exit 1
}
Write-Host 'PASS'
exit 0
