# vybe-test: powershell/type_checking/isnot_operator
if (1 -isnot [string]) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
