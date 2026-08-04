# vybe-test: powershell/scriptblock_invocation/command_in_scriptblock
$script = { Write-Output 'PASS' }
if ((& $script) -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
