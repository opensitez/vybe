# vybe-test: powershell/variable_assignment/assignment_from_property
$obj = [pscustomobject]@{ Name='z' }
$x = $obj.Name
if ($x -eq 'z') { exit 0 }
exit 1
