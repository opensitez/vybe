# vybe-test: powershell/numeric_literals/float_comparison
if (0.1 + 0.2 -ge 0.3) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
