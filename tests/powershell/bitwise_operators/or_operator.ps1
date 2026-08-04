# vybe-test: powershell/bitwise_operators/or_operator
if ((4 -bor 1) -eq 5) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
