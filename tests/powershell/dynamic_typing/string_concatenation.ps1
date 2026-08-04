# vybe-test: powershell/dynamic_typing/string_concatenation
$x = 10
$x = $x + '0'
if ($x -eq '100') { exit 0 }
exit 1
