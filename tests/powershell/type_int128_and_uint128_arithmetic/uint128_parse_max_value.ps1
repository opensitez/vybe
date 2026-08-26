# vybe-test: powershell/type_int128_and_uint128_arithmetic/uint128_parse_max_value
$max = [System.UInt128]::MaxValue
$str = $max.ToString()
if (-not $str.StartsWith("340282366920938463463374607431768211455")) { Write-Host "FAIL: UInt128 MaxValue failed"; exit 1 }
Write-Host "PASS"; exit 0
