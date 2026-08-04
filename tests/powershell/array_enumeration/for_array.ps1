# vybe-test: powershell/array_enumeration/for_array
$arr = 1,2,3
$result = 0
for ($i = 0; $i -lt $arr.Count; $i++) { $result += $arr[$i] }
if ($result -eq 6) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
