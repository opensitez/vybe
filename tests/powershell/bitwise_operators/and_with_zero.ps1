# vybe-test: powershell/bitwise_operators/and_with_zero
if ((5 -band 0) -eq 0) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
