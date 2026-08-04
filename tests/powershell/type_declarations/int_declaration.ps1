# vybe-test: powershell/type_declarations/int_declaration
[int]$x = 5
if ($x -eq 5) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
