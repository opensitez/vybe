# vybe-test: powershell/date_methods/to_string
if ((Get-Date).ToString().Length -gt 0) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
