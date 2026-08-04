# vybe-test: powershell/array_destructuring/array_to_variables_shorter
$a,$b = 1
if ($a -eq 1 -and $b -eq $null) { exit 0 }
exit 1
