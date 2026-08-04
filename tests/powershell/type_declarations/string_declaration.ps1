# vybe-test: powershell/type_declarations/string_declaration
[string]$x = 'PASS'
if ($x -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
