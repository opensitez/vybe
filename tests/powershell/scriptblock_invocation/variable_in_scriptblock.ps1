# vybe-test: powershell/scriptblock_invocation/variable_in_scriptblock
$x = 1
$script = { $x }
if ((& $script) -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
