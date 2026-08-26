# vybe-test: powershell/type_int128_and_uint128_arithmetic/int128_min_max_values
$min = [System.Int128]::MinValue
$max = [System.Int128]::MaxValue
if ($min.CompareTo($max) -ge 0) { Write-Host "FAIL: Int128 Min/Max bounds failed"; exit 1 }
Write-Host "PASS"; exit 0
