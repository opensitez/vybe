# vybe-test: powershell/bitwise_operators/xor_operator
if ((5 -bxor 1) -eq 4) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
