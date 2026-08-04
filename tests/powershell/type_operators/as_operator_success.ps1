# vybe-test: powershell/type_operators/as_operator_success
$value = '123' -as [string]
if ($value -eq '123') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
