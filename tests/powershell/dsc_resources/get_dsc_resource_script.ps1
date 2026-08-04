# vybe-test: powershell/dsc_resources/get_dsc_resource_script
$resource = Get-DscResource -Name Script -ErrorAction SilentlyContinue
if (-not $resource) {
    Write-Host "FAIL: expected Script resource"
    exit 1
}
Write-Host 'PASS'
exit 0
