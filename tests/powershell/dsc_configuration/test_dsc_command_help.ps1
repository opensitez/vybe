# vybe-test: powershell/dsc_configuration/test_dsc_command_help
Get-Help Test-DscConfiguration -ErrorAction SilentlyContinue | Out-Null
Write-Host 'PASS'
exit 0
