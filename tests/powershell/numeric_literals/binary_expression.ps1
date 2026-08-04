# vybe-test: powershell/numeric_literals/binary_expression
if (0b11 * 2 -eq 6) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
