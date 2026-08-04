# vybe-test: powershell/dynamic_typing/hashtable_then_string
$x = @{ a = 1 }
x = 'now'
if ($x -eq 'now') { exit 0 }
exit 1
