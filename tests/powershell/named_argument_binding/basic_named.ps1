# vybe-test: powershell/named_argument_binding/basic_named
function Test { param($x, $y) return "$x,$y" }
if ((Test -y 2 -x 1) -eq '1,2') { exit 0 }
exit 1
