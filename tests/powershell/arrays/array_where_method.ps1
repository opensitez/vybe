# vybe-test: powershell/arrays/array_where_method
$nums = 1..10
$evens = $nums.Where({ $_ % 2 -eq 0 })
if ($evens.Count -ne 5) { Write-Host "FAIL: count"; exit 1 }
if ($evens[0] -ne 2)    { Write-Host "FAIL: first"; exit 1 }
if ($evens[-1] -ne 10)  { Write-Host "FAIL: last"; exit 1 }
Write-Host "PASS"
exit 0
