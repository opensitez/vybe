# vybe-test: powershell/type_int128_and_uint128_arithmetic/uint128_min_value_is_zero
$min = [System.UInt128]::MinValue
if ($min.ToString() -ne "0") { Write-Host "FAIL: UInt128 MinValue expected 0"; exit 1 }
Write-Host "PASS"; exit 0
