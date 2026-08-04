# vybe-test: powershell/array_enumeration/array_member_count
$arr = 1,2,3
if ($arr.Count -eq 3) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
