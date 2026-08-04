# vybe-test: powershell/scriptblock_invocation/inline_scriptblock
$script = { 'PASS' }
if ((& $script) -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
