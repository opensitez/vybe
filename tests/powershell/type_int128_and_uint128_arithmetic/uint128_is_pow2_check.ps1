# vybe-test: powershell/type_int128_and_uint128_arithmetic/uint128_is_pow2_check
$pow = [System.UInt128]::Parse("1024")
$notPow = [System.UInt128]::Parse("1025")
if (-not [System.UInt128]::IsPow2($pow) -or [System.UInt128]::IsPow2($notPow)) { Write-Host "FAIL: UInt128 IsPow2 failed"; exit 1 }
Write-Host "PASS"; exit 0
