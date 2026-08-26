# vybe-test: powershell/type_int128_and_uint128_arithmetic/int128_parse_large_number
$val = [System.Int128]::Parse("34028236692093846346337460743176821145")
$str = $val.ToString()
if ($str -ne "34028236692093846346337460743176821145") { Write-Host "FAIL: Int128 Parse failed"; exit 1 }
Write-Host "PASS"; exit 0
