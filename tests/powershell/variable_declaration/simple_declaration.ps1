# vybe-test: powershell/variable_declaration/simple_declaration
$x = 1
if ($x -ne 1) { Write-Host 'FAIL'; exit 1 }
Write-Host 'PASS'
exit 0
