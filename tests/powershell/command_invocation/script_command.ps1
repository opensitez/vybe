# vybe-test: powershell/command_invocation/script_command
$script = { 'PASS' }
if ((& $script) -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
