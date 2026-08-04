# vybe-test: powershell/expression_evaluation/assignment_expression
$x = 1 + 2
if ($x -eq 3) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
