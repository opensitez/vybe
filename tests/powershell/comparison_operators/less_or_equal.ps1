# vybe-test: powershell/comparison_operators/less_or_equal
if (2 -le 2) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
