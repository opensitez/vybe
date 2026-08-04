# vybe-test: powershell/script_scoping/explicit_scope
$script:a = 'outer'
function Test { $script:a = 'inner' }
Test
if ($script:a -eq 'inner') { exit 0 }
exit 1
