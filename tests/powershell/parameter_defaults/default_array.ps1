# vybe-test: powershell/parameter_defaults/default_array
function Test { param($x = 1,2,3) return $x }
if ((Test).Count -eq 3) { exit 0 }
exit 1
