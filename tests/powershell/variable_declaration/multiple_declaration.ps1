# vybe-test: powershell/variable_declaration/multiple_declaration
$x = $y = 1
if ($x -ne 1 -or $y -ne 1) { Write-Host 'FAIL'; exit 1 }
Write-Host 'PASS'
exit 0
