# vybe-test: powershell/type_constructors/regex_constructor
$value = [regex]'a'
if ($value -eq $null) { Write-Host 'FAIL'; exit 1 }
Write-Host 'PASS'
exit 0
