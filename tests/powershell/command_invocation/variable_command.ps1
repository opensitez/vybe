# vybe-test: powershell/command_invocation/variable_command
$cmd = 'Write-Output'
if ((& $cmd 'PASS') -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
