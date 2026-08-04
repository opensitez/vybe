# vybe-test: powershell/script_invocation/invoke_expression
$expr = '1 + 2'
if ((Invoke-Expression $expr) -eq 3) { exit 0 }
exit 1
