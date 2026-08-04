# vybe-test: powershell/variable_assignment/assignment_complex_expression
$x = 1; $x = $x * 2
if ($x -eq 2) { exit 0 }
exit 1
