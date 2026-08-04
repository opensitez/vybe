# vybe-test: powershell/property_access/property_on_command_result
if ((Get-Date).Year -gt 2000) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
