# vybe-test: powershell/date_methods/day_of_year
if ((Get-Date).DayOfYear -gt 0) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
