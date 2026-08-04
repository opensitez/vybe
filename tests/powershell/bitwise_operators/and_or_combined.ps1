# vybe-test: powershell/bitwise_operators/and_or_combined
if (((5 -band 1) -bor 2) -eq 3) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
