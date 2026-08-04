# vybe-test: powershell/type_checking/as_operator_failure
$value = '123' -as [int]
if ($value -eq $null) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
