# vybe-test: powershell/named_argument_binding/mixed_named_positional
function Test { param($x, $y, $z) return "$x,$y,$z" }
if ((Test 1 -z 3 -y 2) -eq '1,2,3') { exit 0 }
exit 1
