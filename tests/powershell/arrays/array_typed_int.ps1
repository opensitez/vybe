# vybe-test: powershell/arrays/array_typed_int
[int[]]$nums = 1, 2, 3
$nums += 4
if ($nums.Count -ne 4)  { Write-Host "FAIL: count"; exit 1 }
if ($nums[-1] -ne 4)    { Write-Host "FAIL: last";  exit 1 }
# Typed array rejects wrong type via coercion
[int[]]$coerced = "10", "20"
if ($coerced[0] -ne 10) { Write-Host "FAIL: coercion"; exit 1 }
Write-Host "PASS"
exit 0
