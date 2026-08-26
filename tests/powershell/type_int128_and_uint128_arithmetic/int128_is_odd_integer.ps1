# vybe-test: powershell/type_int128_and_uint128_arithmetic/int128_is_odd_integer
$odd = [System.Int128]::Parse("101")
$even = [System.Int128]::Parse("100")
if (-not [System.Int128]::IsOddInteger($odd) -or [System.Int128]::IsOddInteger($even)) { Write-Host "FAIL: Int128 IsOddInteger failed"; exit 1 }
Write-Host "PASS"; exit 0
