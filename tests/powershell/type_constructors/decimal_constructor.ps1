# vybe-test: powershell/type_constructors/decimal_constructor
$value = [decimal]'1.5'
if ($value -ne 1.5) { Write-Host 'FAIL'; exit 1 }
Write-Host 'PASS'
exit 0
