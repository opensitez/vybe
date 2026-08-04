# vybe-test: powershell/type_constructors/object_constructor
$value = [object]'hello'
if ($value -ne 'hello') { Write-Host 'FAIL'; exit 1 }
Write-Host 'PASS'
exit 0
