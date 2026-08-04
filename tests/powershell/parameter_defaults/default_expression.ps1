# vybe-test: powershell/parameter_defaults/default_expression
function Test { param($x = 1 + 1) return $x }
if ((Test) -eq 2) { exit 0 }
exit 1
