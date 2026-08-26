# vybe-test: powershell/type_int128_and_uint128_arithmetic/int128_hashcode_consistency
$a = [System.Int128]::Parse("77777777")
$b = [System.Int128]::Parse("77777777")
if ($a.GetHashCode() -ne $b.GetHashCode()) { Write-Host "FAIL: Int128 HashCode failed"; exit 1 }
Write-Host "PASS"; exit 0
