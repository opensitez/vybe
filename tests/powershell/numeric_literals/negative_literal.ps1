# vybe-test: powershell/numeric_literals/negative_literal
if (-42 -eq -42) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
