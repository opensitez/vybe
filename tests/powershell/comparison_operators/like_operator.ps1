# vybe-test: powershell/comparison_operators/like_operator
if ('hello' -like 'h*') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
