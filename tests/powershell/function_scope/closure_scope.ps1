# vybe-test: powershell/function_scope/closure_scope
$x = 1
$sb = { $x }
if (($sb.Invoke()) -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
