# vybe-test: powershell/bitwise_rotation_and_shifts/leading_zero_count
[uint32]$val = 0x00000001
$lz = [System.Numerics.BitOperations]::LeadingZeroCount($val)
if ($lz -ne 31) {
    Write-Host "FAIL: LeadingZeroCount expected 31, got $lz"
    exit 1
}
Write-Host "PASS"
exit 0
