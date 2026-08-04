# vybe-test: powershell/script_scoping/global_assignment
function Test { $global:a = 'x' }
Test
if ($a -eq 'x') { exit 0 }
exit 1
