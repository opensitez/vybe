# vybe-test: powershell/scope_resolution/scriptblock_scope
$x = 1
$sb = { $x = 2 }
& $sb
if ($x -eq 2) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
