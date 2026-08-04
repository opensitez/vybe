# vybe-test: powershell/named_argument_binding/partial_named
function Test { param($x, $y) return "$x,$y" }
if ((Test 1 -y 2) -eq '1,2') { exit 0 }
exit 1
