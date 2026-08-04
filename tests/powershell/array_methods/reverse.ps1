# vybe-test: powershell/array_methods/reverse
$arr = 1,2,3
[Array]::Reverse($arr)
if ($arr[0] -eq 3) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
