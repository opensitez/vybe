# vybe-test: powershell/type_declarations/typed_variable_assignment
[int]$x = 1
$x = 2
if ($x -eq 2) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
