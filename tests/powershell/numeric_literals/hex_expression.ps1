# vybe-test: powershell/numeric_literals/hex_expression
if (0xA + 5 -eq 15) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
