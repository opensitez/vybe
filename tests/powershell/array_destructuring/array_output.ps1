# vybe-test: powershell/array_destructuring/array_output
$a,$b = 9,8
if ((,$a,$b).Count -eq 3) { exit 0 }
exit 1
