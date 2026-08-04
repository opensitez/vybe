# vybe-test: powershell/bitwise_operators/right_shift
if ((4 -shr 1) -eq 2) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
