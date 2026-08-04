# vybe-test: powershell/script_scoping/function_scope
function Test { $a = 'x' }
Test
if ($a -eq $null) { exit 0 }
exit 1
