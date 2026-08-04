# vybe-test: powershell/variable_assignment/assignment_null
$x = $null
if ($x -eq $null) { exit 0 }
exit 1
