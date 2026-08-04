# vybe-test: powershell/named_argument_binding/pipeline_named
function Test { param($x, $y) return "$x,$y" }
if ((Test -x 1 -y 2) -eq '1,2') { exit 0 }
exit 1
