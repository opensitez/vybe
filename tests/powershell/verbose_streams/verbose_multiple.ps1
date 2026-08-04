# vybe-test: powershell/verbose_streams/verbose_multiple
$VerbosePreference = 'Continue'
Write-Verbose 'first'
Write-Verbose 'second'
Write-Host 'PASS'
exit 0
