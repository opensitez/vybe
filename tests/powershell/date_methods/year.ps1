# vybe-test: powershell/date_methods/year
if ((Get-Date).Year -ge 2000) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
