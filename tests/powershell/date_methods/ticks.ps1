# vybe-test: powershell/date_methods/ticks
if ((Get-Date).Ticks -gt 0) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
