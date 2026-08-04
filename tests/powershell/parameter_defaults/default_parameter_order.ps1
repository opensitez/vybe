# vybe-test: powershell/parameter_defaults/default_parameter_order
function Test { param($x = 1, $y = 2) return "$x,$y" }
if ((Test) -eq '1,2') { exit 0 }
exit 1
