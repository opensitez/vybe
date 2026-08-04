# vybe-test: powershell/bitwise_operators/left_shift
if ((1 -shl 2) -eq 4) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
