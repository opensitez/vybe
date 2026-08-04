# vybe-test: powershell/verbose_streams/verbose_command
$VerbosePreference = 'Continue'
Write-Verbose 'test'
Write-Host 'PASS'
exit 0
