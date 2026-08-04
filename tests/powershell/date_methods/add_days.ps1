# vybe-test: powershell/date_methods/add_days
if ((Get-Date).AddDays(1).Day -ne (Get-Date).Day) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
