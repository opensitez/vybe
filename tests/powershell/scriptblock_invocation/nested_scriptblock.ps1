# vybe-test: powershell/scriptblock_invocation/nested_scriptblock
$inner = { 'PASS' }
$outer = { & $inner }
if ((& $outer) -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
