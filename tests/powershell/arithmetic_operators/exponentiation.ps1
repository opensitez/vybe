# vybe-test: powershell/arithmetic_operators/exponentiation
if ((2 ** 3) -ne 8) { Write-Host 'FAIL'; exit 1 }
Write-Host 'PASS'
exit 0
