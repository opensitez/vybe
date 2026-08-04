# vybe-test: powershell/date_methods/today
if ((Get-Date).Date -eq (Get-Date).Date) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
