# vybe-test: powershell/dynamic_typing/implicit_array
$x = 1,2
if ($x.Count -eq 2) { exit 0 }
exit 1
