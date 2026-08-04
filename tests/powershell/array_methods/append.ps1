# vybe-test: powershell/array_methods/append
$arr = 1,2
$arr += 3
if ($arr[-1] -eq 3) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
