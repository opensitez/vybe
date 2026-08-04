# vybe-test: powershell/script_scoping/child_scope
function Test { $script:a = 3 }
Test
if ($script:a -eq 3) { exit 0 }
exit 1
