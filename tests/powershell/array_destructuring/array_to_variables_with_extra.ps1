# vybe-test: powershell/array_destructuring/array_to_variables_with_extra
$a,$b = 1,2,3
if ($a -eq 1 -and $b -eq 2) { exit 0 }
exit 1
