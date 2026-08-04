# vybe-test: powershell/variable_declaration/variable_reassignment
$x = 1
$x = 2
if ($x -ne 2) { Write-Host 'FAIL'; exit 1 }
Write-Host 'PASS'
exit 0
