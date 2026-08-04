# vybe-test: powershell/type_operators/is_operator
if (1 -is [int]) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
