# vybe-test: powershell/dynamic_typing/reassign_integer_string
$x = 1
$x = 'one'
if ($x -eq 'one') { exit 0 }
exit 1
