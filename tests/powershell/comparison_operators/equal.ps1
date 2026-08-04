# vybe-test: powershell/comparison_operators/equal
if (1 -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
