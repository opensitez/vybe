# vybe-test: powershell/arithmetic_operators/unary_minus
if ((-5) -ne -5) { Write-Host 'FAIL'; exit 1 }
Write-Host 'PASS'
exit 0
