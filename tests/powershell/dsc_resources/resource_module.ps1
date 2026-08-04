# vybe-test: powershell/dsc_resources/resource_module
$resource = Get-DscResource -Name File -ErrorAction SilentlyContinue
if ($resource.ModuleName -eq $null) {
    Write-Host "FAIL: expected module name"
    exit 1
}
Write-Host 'PASS'
exit 0
