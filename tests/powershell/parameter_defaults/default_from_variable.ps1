# vybe-test: powershell/parameter_defaults/default_from_variable
$val = 5
function Test { param($x = $val) return $x }
if ((Test) -eq 5) { exit 0 }
exit 1
