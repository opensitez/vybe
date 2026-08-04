# vybe-test: powershell/dsc_resources/resource_names
$names = (Get-DscResource -ErrorAction SilentlyContinue).Name
if ($names -notcontains 'File') {
    Write-Host "FAIL: expected File resource name"
    exit 1
}
Write-Host 'PASS'
exit 0
