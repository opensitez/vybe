# vybe-test: powershell/string_literal_quotes/literal_variable
$name = 'PASS'
if ('$name' -eq '$name') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
