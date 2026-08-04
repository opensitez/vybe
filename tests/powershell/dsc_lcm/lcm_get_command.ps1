# vybe-test: powershell/dsc_lcm/lcm_get_command
$cmd = Get-Command Get-DscLocalConfigurationManager -ErrorAction SilentlyContinue
if (-not $cmd) {
    Write-Host "FAIL"
    exit 1
}
Write-Host 'PASS'
exit 0
