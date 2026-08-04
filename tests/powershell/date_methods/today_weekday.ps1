# vybe-test: powershell/date_methods/today_weekday
if ((Get-Date).DayOfWeek -ne $null) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
