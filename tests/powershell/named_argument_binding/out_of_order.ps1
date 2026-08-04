# vybe-test: powershell/named_argument_binding/out_of_order
function Test { param($x, $y) return "$x,$y" }
if ((Test -y 'b' -x 'a') -eq 'a,b') { exit 0 }
exit 1
