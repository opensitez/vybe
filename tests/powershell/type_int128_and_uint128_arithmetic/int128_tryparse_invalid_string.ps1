# vybe-test: powershell/type_int128_and_uint128_arithmetic/int128_tryparse_invalid_string
$outVal = [System.Int128]::Zero
$ok = [System.Int128]::TryParse("invalid_int128", [ref]$outVal)
if ($ok) { Write-Host "FAIL: Int128 TryParse should fail on invalid string"; exit 1 }
Write-Host "PASS"; exit 0
