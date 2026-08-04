# vybe-test: powershell/arithmetic_operators/subtraction
if ((5 - 2) -ne 3) { Write-Host 'FAIL'; exit 1 }
Write-Host 'PASS'
exit 0
