# vybe-test: powershell/bitwise_operators/or_with_zero
if ((5 -bor 0) -eq 5) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
