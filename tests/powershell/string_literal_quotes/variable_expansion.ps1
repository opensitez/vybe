# vybe-test: powershell/string_literal_quotes/variable_expansion
$name = 'PASS'
if ("$name" -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
