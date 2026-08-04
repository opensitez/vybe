# vybe-test: powershell/arithmetic_operators/negative_number
if ((-2 + 3) -ne 1) { Write-Host 'FAIL'; exit 1 }
Write-Host 'PASS'
exit 0
