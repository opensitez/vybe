# vybe-test: powershell/comparison_operators/greater_than
if (3 -gt 2) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
