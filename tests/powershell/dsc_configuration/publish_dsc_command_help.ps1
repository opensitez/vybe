# vybe-test: powershell/dsc_configuration/publish_dsc_command_help.ps1
Get-Help Publish-DscConfiguration -ErrorAction SilentlyContinue | Out-Null
Write-Host 'PASS'
exit 0
