# vybe-test: powershell/comparison_operators/greater_or_equal
if (2 -ge 2) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
