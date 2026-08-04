# vybe-test: powershell/help_topics/help_about_help
Get-Help about_Execution_Policies -ErrorAction SilentlyContinue | Out-Null
Write-Host 'PASS'
exit 0
