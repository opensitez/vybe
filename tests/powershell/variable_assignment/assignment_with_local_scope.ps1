# vybe-test: powershell/variable_assignment/assignment_with_local_scope
function Test { Set-Variable -Name x -Value 11 -Scope Local; return $x }
if ((Test) -eq 11) { exit 0 }
exit 1
