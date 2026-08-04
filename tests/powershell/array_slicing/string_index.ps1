# vybe-test: powershell/array_slicing/string_index
$str = 'abc'
if ($str[1] -eq 'b') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
