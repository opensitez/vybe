# vybe-test: powershell/arithmetic_operators/division
if ((6 / 2) -ne 3) { Write-Host 'FAIL'; exit 1 }
Write-Host 'PASS'
exit 0
