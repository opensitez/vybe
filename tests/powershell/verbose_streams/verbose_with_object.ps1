# vybe-test: powershell/verbose_streams/verbose_with_object
$VerbosePreference = 'Continue'
Write-Verbose ([PSCustomObject]@{ A='x' })
Write-Host 'PASS'
exit 0
