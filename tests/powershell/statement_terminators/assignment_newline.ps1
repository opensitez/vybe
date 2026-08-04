# vybe-test: powershell/statement_terminators/assignment_newline
$x =
1
if ($x -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
