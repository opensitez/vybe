# vybe-test: powershell/math_clamp_and_range_boundaries/int64_min_and_max
[int64]$a = 9000000000000000000
[int64]$b = 9000000000000000001
$m = [math]::Max($a, $b)
if ($m -ne $b) {
    Write-Host "FAIL: [math]::Max on int64 failed"
    exit 1
}
Write-Host "PASS"
exit 0
