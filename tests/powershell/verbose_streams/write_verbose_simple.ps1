# vybe-test: powershell/verbose_streams/write_verbose_simple
$VerbosePreference = 'Continue'
Write-Verbose 'verb'
Write-Host 'PASS'
exit 0
