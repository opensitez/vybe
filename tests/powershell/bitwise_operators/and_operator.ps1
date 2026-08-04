# vybe-test: powershell/bitwise_operators/and_operator
if ((5 -band 1) -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
