# vybe-test: powershell/type_int128_and_uint128_arithmetic/int128_zero_and_one_constants
$z = [System.Int128]::Zero
$o = [System.Int128]::One
if ($z.ToString() -ne "0" -or $o.ToString() -ne "1") { Write-Host "FAIL: Int128 Zero/One constants failed"; exit 1 }
Write-Host "PASS"; exit 0
