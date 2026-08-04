# vybe-test: powershell/comparison_operators/match_operator
if ('hello' -match 'he') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
