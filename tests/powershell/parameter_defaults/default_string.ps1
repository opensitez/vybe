# vybe-test: powershell/parameter_defaults/default_string
function Test { param($x = 'hello') return $x }
if ((Test) -eq 'hello') { exit 0 }
exit 1
