# vybe-test: powershell/command_subexpressions/string_subexpression
$name = 'PASS'
if ("Hello $($name)" -eq 'Hello PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
