# vybe-test: powershell/type_int128_and_uint128_arithmetic/uint128_equality_check
$u1 = [System.UInt128]::Parse("99999999999999999999")
$u2 = [System.UInt128]::Parse("99999999999999999999")
if (-not $u1.Equals($u2)) { Write-Host "FAIL: UInt128 Equals failed"; exit 1 }
Write-Host "PASS"; exit 0
