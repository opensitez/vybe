# vybe-test: powershell/scriptblock_invocation/scriptblock_returning_object
$script = { [pscustomobject]@{ A=1 } }
if ((& $script).A -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
