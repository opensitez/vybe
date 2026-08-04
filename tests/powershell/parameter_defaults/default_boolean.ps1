# vybe-test: powershell/parameter_defaults/default_boolean
function Test { param($x = $true) return $x }
if ((Test) -eq $true) { exit 0 }
exit 1
