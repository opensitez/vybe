# vybe-test: powershell/scriptblock_invocation/scriptblock_with_param
$script = { param($x) $x }
if ((& $script 5) -eq 5) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
