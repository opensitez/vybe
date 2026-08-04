# vybe-test: powershell/array_destructuring/array_assignment
$x = 1,2,3
$a,$b,$c = $x
if ($a -eq 1 -and $b -eq 2 -and $c -eq 3) { exit 0 }
exit 1
