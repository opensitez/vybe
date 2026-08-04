# vybe-test: powershell/scriptblock_invocation/defined_scriptblock
function Test-Func { $sb = { 'PASS' }; (& $sb) }
if ((Test-Func) -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
