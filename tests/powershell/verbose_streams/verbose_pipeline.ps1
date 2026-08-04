# vybe-test: powershell/verbose_streams/verbose_pipeline
$VerbosePreference = 'Continue'
1..2 | ForEach-Object { Write-Verbose 'v' }
Write-Host 'PASS'
exit 0
