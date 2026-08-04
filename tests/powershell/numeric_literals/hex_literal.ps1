# vybe-test: powershell/numeric_literals/hex_literal
if (0xF -eq 15) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
