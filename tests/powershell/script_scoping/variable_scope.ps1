# vybe-test: powershell/script_scoping/variable_scope
$global:a = 1
if ($a -eq 1) { exit 0 }
exit 1
