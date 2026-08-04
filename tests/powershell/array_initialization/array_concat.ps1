# vybe-test: powershell/array_initialization/array_concat
$arr = 1,2 + 3,4
if ($arr.Count -eq 4) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
