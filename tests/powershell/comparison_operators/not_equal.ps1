# vybe-test: powershell/comparison_operators/not_equal
if (1 -ne 2) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
