# vybe-test: powershell/string_methods/length
if ('hello'.Length -eq 5) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
