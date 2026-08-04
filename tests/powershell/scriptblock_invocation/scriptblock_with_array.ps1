# vybe-test: powershell/scriptblock_invocation/scriptblock_with_array
$script = { @(1,2,3) }
if ((& $script).Count -eq 3) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
