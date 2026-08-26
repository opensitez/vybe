# vybe-test: powershell/type_int128_and_uint128_arithmetic/int128_abs_method
$neg = [System.Int128]::Parse("-123456789")
$abs = [System.Int128]::Abs($neg)
if ($abs.ToString() -ne "123456789") { Write-Host "FAIL: Int128 Abs failed"; exit 1 }
Write-Host "PASS"; exit 0
