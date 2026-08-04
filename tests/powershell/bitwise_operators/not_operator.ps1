# vybe-test: powershell/bitwise_operators/not_operator
if ((5 -bnot 5) -eq -6) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
