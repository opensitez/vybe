# vybe-test: powershell/bitwise_rotation_and_shifts/trailing_zero_count
[uint32]$val = 0x00000008 # bit 3 set -> 3 trailing zeros
$tz = [System.Numerics.BitOperations]::TrailingZeroCount($val)
if ($tz -ne 3) {
    Write-Host "FAIL: TrailingZeroCount expected 3, got $tz"
    exit 1
}
Write-Host "PASS"
exit 0
