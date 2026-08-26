# vybe-test: powershell/bitwise_leading_trailing_zero_count/leading_trailing_zero_case_15
$lz = [System.Numerics.BitOperations]::LeadingZeroCount([uint32]1)
$tz = [System.Numerics.BitOperations]::TrailingZeroCount([uint32]8)
if ($lz -ne 31 -or $tz -ne 3) { Write-Host "FAIL: Leading/Trailing zero count failed"; exit 1 }
Write-Host "PASS"; exit 0
