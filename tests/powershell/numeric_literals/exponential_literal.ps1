# vybe-test: powershell/numeric_literals/exponential_literal
if (1e3 -eq 1000) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
