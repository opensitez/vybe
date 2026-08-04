# vybe-test: powershell/property_access/date_property
if ((Get-Date).Day -ge 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
