# vybe-test: powershell/bitwise_operators/xor_self
if ((5 -bxor 5) -eq 0) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
