# vybe-test: powershell/comparison_operators/not_in_operator
if (4 -notin (1,2,3)) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
