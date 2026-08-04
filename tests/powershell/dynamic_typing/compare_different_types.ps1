# vybe-test: powershell/dynamic_typing/compare_different_types
$x = '1'
$y = 1
if ($x -ne $y) { exit 0 }
exit 1
