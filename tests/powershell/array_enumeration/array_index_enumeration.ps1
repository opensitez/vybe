# vybe-test: powershell/array_enumeration/array_index_enumeration
$arr = 1,2,3
$result = ''
foreach ($i in 0..($arr.Count - 1)) { $result += $arr[$i] }
if ($result -eq '123') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
