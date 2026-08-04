# vybe-test: powershell/type_constructors/hashtable_constructor
$value = [hashtable]@{ a=1 }
if ($value.a -ne 1) { Write-Host 'FAIL'; exit 1 }
Write-Host 'PASS'
exit 0
