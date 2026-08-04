# vybe-test: powershell/named_argument_binding/explicit_switch
function Test { param($x, $y) return $y }
if ((Test -x 1 -y 2) -eq 2) { exit 0 }
exit 1
