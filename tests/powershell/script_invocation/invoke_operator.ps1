# vybe-test: powershell/script_invocation/invoke_operator
$cmd = { Write-Output 5 }
if ((& $cmd) -eq 5) { exit 0 }
exit 1
