# vybe-test: powershell/dsc_lcm/get_lcm_command
if (-not (Get-Command Get-DscLocalConfigurationManager -ErrorAction SilentlyContinue)) {
    Write-Host "FAIL: Get-DscLocalConfigurationManager unavailable"
    exit 1
}
Write-Host 'PASS'
exit 0
