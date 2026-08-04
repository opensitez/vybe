# vybe-test: powershell/arithmetic_operators/addition
if ((1 + 2) -ne 3) { Write-Host 'FAIL'; exit 1 }
Write-Host 'PASS'
exit 0
