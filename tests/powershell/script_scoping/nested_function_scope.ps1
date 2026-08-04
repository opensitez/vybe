# vybe-test: powershell/script_scoping/nested_function_scope
function Outer { function Inner { $a = 'inner' }; Inner }
Outer
if ($a -eq $null) { exit 0 }
exit 1
