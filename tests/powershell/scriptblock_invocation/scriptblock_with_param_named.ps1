# vybe-test: powershell/scriptblock_invocation/scriptblock_with_param_named
$script = { param($x) $x }
if ((& $script -x 9) -eq 9) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
