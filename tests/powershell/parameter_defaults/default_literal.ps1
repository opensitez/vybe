# vybe-test: powershell/parameter_defaults/default_literal
function Test { param($x = 1) return $x }
if ((Test) -eq 1) { exit 0 }
exit 1
