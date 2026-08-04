# vybe-test: powershell/type_operators/type_operator_expression
if ((1 -is [int]) -and ('x' -is [string])) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
