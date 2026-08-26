# vybe-test: powershell/type_int128_and_uint128_arithmetic/uint128_compareto_ordering
$u1 = [System.UInt128]::Parse("1000")
$u2 = [System.UInt128]::Parse("2000")
if ($u1.CompareTo($u2) -ge 0 -or $u2.CompareTo($u1) -le 0) { Write-Host "FAIL: UInt128 CompareTo failed"; exit 1 }
Write-Host "PASS"; exit 0
