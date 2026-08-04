# vybe-test: powershell/script_scoping/locale_scope
function Test { $a = 2; return $a }
if ((Test) -eq 2) { exit 0 }
exit 1
