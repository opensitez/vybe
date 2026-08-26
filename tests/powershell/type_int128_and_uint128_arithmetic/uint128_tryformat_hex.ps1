# vybe-test: powershell/type_int128_and_uint128_arithmetic/uint128_tryformat_hex
$u = [System.UInt128]::Parse("255")
$hex = $u.ToString("X")
if ($hex -ne "FF") { Write-Host "FAIL: UInt128 hex format failed, got $hex"; exit 1 }
Write-Host "PASS"; exit 0
