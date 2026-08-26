# vybe-test: powershell/numeric_random_number_generation/next_int64_range_bounds
$rnd = [System.Random]::new()
$min = 1000000000000
$max = 2000000000000
$v = $rnd.NextInt64($min, $max)
if ($v -lt $min -or $v -ge $max) {
    Write-Host "FAIL: NextInt64 out of range: $v"
    exit 1
}
Write-Host "PASS"
exit 0
