# vybe-test: powershell/parameter_defaults/default_hashtable
function Test { param($x = @{ a=1 }) return $x }
if ((Test).a -eq 1) { exit 0 }
exit 1
