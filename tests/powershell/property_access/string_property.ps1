# vybe-test: powershell/property_access/string_property
if ('hello'.Length -eq 5) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
