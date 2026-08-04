# vybe-test: powershell/type_declarations/bool_declaration
[bool]$x = $true
if ($x) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
