# vybe-test: powershell/dsc_lcm/lcm_command_exists
if (-not (Get-Command Get-DscLocalConfigurationManager -ErrorAction SilentlyContinue)) {
    Write-Host "FAIL: missing command"
    exit 1
}
Write-Host 'PASS'
exit 0
