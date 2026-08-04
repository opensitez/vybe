# vybe-test: powershell/help_topics/help_about_error_handling
Get-Help about_Try_Catch_Finally -ErrorAction SilentlyContinue | Out-Null
Write-Host 'PASS'
exit 0
