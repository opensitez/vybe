# vybe-test: powershell/named_argument_binding/alias_name
function Test { param([Alias('a')]$x) return $x }
if ((Test -a 5) -eq 5) { exit 0 }
exit 1
