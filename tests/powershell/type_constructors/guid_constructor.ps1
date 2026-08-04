# vybe-test: powershell/type_constructors/guid_constructor
$value = [guid]'00000000-0000-0000-0000-000000000000'
if ($value -eq $null) { Write-Host 'FAIL'; exit 1 }
Write-Host 'PASS'
exit 0
