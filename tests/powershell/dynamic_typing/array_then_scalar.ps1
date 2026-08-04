# vybe-test: powershell/dynamic_typing/array_then_scalar
$x = @(1,2,3)
x = 4
if ($x -eq 4) { exit 0 }
exit 1
