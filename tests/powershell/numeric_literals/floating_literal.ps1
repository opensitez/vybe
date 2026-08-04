# vybe-test: powershell/numeric_literals/floating_literal
if (1.5 -eq 1.5) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
