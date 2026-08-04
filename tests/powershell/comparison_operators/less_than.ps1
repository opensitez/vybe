# vybe-test: powershell/comparison_operators/less_than
if (2 -lt 3) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
