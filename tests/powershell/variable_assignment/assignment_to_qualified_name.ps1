# vybe-test: powershell/variable_assignment/assignment_to_qualified_name
$script:x = 10
if ($script:x -eq 10) { exit 0 }
exit 1
