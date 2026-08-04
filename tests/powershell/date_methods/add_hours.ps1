# vybe-test: powershell/date_methods/add_hours
if ((Get-Date).AddHours(1).Hour -ne (Get-Date).Hour) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
