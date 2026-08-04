# vybe-test: powershell/command_invocation/call_operator
if ((& { 'PASS' }) -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
