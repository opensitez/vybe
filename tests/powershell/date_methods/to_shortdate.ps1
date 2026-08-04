# vybe-test: powershell/date_methods/to_shortdate
if ((Get-Date).ToShortDateString() -ne '') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
