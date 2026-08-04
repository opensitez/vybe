# vybe-test: powershell/numeric_literals/binary_literal
if (0b1010 -eq 10) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
