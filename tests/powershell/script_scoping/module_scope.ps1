# vybe-test: powershell/script_scoping/module_scope
$script:a = 'x'
if ($script:a -eq 'x') { exit 0 }
exit 1
