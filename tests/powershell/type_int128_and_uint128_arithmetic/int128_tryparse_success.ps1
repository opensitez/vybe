# vybe-test: powershell/type_int128_and_uint128_arithmetic/int128_tryparse_success
$outVal = [System.Int128]::Zero
$ok = [System.Int128]::TryParse("12345678901234567890", [ref]$outVal)
if (-not $ok -or $outVal.ToString() -ne "12345678901234567890") { Write-Host "FAIL: Int128 TryParse failed"; exit 1 }
Write-Host "PASS"; exit 0
