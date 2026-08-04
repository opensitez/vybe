# vybe-test: powershell/variable_assignment/assignment_after_loop
for ($i=0; $i -lt 1; $i++) { $x = 9 }
if ($x -eq 9) { exit 0 }
exit 1
