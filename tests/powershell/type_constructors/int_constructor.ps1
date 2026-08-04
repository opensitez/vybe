# vybe-test: powershell/type_constructors/int_constructor
$value = [int]'5'
if ($value -ne 5) { Write-Host 'FAIL'; exit 1 }
Write-Host 'PASS'
exit 0
