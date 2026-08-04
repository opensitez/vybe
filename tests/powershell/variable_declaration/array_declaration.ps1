# vybe-test: powershell/variable_declaration/array_declaration
$arr = 1,2,3
if ($arr.Length -ne 3) { Write-Host 'FAIL'; exit 1 }
Write-Host 'PASS'
exit 0
