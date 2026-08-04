# vybe-test: powershell/dsc_resources/resource_help
$resource = Get-DscResource -Name File -ErrorAction SilentlyContinue
if ($resource -eq $null) {
    Write-Host "FAIL: expected File resource"
    exit 1
}
Write-Host 'PASS'
exit 0
