# vybe-test: powershell/date_methods/add_minutes
if ((Get-Date).AddMinutes(1).Minute -ne (Get-Date).Minute) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
