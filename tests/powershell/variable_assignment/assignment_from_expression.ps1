# vybe-test: powershell/variable_assignment/assignment_from_expression
$x = 1 + 2
if ($x -eq 3) { exit 0 }
exit 1
