# vybe-test: powershell/array_enumeration/array_enumerate_empty
$arr = @()
$result = 0
foreach ($x in $arr) { $result++ }
if ($result -eq 0) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
