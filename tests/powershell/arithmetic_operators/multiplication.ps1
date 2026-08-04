# vybe-test: powershell/arithmetic_operators/multiplication
if ((3 * 2) -ne 6) { Write-Host 'FAIL'; exit 1 }
Write-Host 'PASS'
exit 0
