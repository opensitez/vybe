# vybe-test: powershell/array_enumeration/array_enumerate_strings
$arr = 'a','b','c'
if (($arr | ForEach-Object { $_.ToUpper() }) -join ',' -eq 'A,B,C') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
