# vybe-test: powershell/array_destructuring/array_to_variables
$a,$b = 1,2
if ($a -eq 1 -and $b -eq 2) { exit 0 }
exit 1
