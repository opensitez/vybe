# vybe-test: powershell/type_int128_and_uint128_arithmetic/int128_negative_one_constant
$negOne = [System.Int128]::NegativeOne
if ($negOne.ToString() -ne "-1") { Write-Host "FAIL: Int128 NegativeOne failed"; exit 1 }
Write-Host "PASS"; exit 0
