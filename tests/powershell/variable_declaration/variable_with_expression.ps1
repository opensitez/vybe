# vybe-test: powershell/variable_declaration/variable_with_expression
$x = 1 + 1
if ($x -ne 2) { Write-Host 'FAIL'; exit 1 }
Write-Host 'PASS'
exit 0
