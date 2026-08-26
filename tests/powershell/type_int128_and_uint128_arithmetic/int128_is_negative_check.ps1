# vybe-test: powershell/type_int128_and_uint128_arithmetic/int128_is_negative_check
$neg = [System.Int128]::Parse("-1")
$pos = [System.Int128]::Parse("1")
if (-not [System.Int128]::IsNegative($neg) -or [System.Int128]::IsNegative($pos)) { Write-Host "FAIL: Int128 IsNegative failed"; exit 1 }
Write-Host "PASS"; exit 0
