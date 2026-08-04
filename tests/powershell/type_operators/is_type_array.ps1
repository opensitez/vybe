# vybe-test: powershell/type_operators/is_type_array
if ((1,2,3) -is [object[]]) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
