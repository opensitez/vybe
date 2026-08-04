# vybe-test: powershell/dsc_lcm/lcm_query
$cmd = Get-Command Get-DscLocalConfigurationManager -ErrorAction SilentlyContinue
if (-not $cmd) {
    Write-Host "FAIL: unavailable"
    exit 1
}
Write-Host 'PASS'
exit 0
