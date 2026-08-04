# vybe-test: powershell/dsc_lcm/lcm_command_help
Get-Help Get-DscLocalConfigurationManager -ErrorAction SilentlyContinue | Out-Null
Write-Host 'PASS'
exit 0
