# vybe-test: powershell/dsc_lcm/lcm_error_handling
if (-not (Get-Command Get-DscLocalConfigurationManager -ErrorAction SilentlyContinue)) {
    Write-Host "FAIL"
    exit 1
}
Write-Host 'PASS'
exit 0
