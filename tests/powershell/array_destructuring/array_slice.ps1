# vybe-test: powershell/array_destructuring/array_slice
$x = 1,2,3,4
$y = $x[1..2]
if ($y[0] -eq 2 -and $y[1] -eq 3) { exit 0 }
exit 1
