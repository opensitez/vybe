# vybe-test: powershell/bitwise_log2_and_power_of_two/log2_and_is_power_case_11
$log2 = [System.Numerics.BitOperations]::Log2([uint32]64)
$isPow = [System.Numerics.BitOperations]::IsPow2([uint32]64)
if ($log2 -ne 6 -or -not $isPow) { Write-Host "FAIL: Log2 / IsPow2 failed"; exit 1 }
Write-Host "PASS"; exit 0
