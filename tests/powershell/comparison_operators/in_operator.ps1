# vybe-test: powershell/comparison_operators/in_operator
if (1 -in (1,2,3)) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
