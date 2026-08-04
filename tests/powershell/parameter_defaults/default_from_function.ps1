# vybe-test: powershell/parameter_defaults/default_from_function
function DefaultVal { 7 }
function Test { param($x = (DefaultVal)) return $x }
if ((Test) -eq 7) { exit 0 }
exit 1
