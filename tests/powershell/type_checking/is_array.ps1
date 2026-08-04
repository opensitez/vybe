# vybe-test: powershell/type_checking/is_array
if ((1,2,3) -is [object[]]) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
