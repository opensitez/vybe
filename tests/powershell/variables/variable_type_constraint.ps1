# vybe-test: powershell/variables/variable_type_constraint
[int]$count = 0
$count = 5
if ($count -ne 5) { Write-Host "FAIL: assignment"; exit 1 }
# String assigned to typed int gets coerced
$count = "10"
if ($count -ne 10) { Write-Host "FAIL: coerce string->int"; exit 1 }
if ($count -isnot [int]) { Write-Host "FAIL: type should still be int"; exit 1 }
Write-Host "PASS"
exit 0
