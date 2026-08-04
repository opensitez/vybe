# vybe-test: powershell/scriptblock_invocation/scriptblock_in_pipe
$script = { $_ }
if ((1 | ForEach-Object $script) -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
