# vybe-test: powershell/type_int128_and_uint128_arithmetic/int128_is_even_integer
$even = [System.Int128]::Parse("100")
$odd = [System.Int128]::Parse("101")
if (-not [System.Int128]::IsEvenInteger($even) -or [System.Int128]::IsEvenInteger($odd)) { Write-Host "FAIL: Int128 IsEvenInteger failed"; exit 1 }
Write-Host "PASS"; exit 0
