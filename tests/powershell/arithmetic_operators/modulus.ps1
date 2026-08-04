# vybe-test: powershell/arithmetic_operators/modulus
if ((7 % 3) -ne 1) { Write-Host 'FAIL'; exit 1 }
Write-Host 'PASS'
exit 0
