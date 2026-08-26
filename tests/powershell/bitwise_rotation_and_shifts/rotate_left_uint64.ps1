# vybe-test: powershell/bitwise_rotation_and_shifts/rotate_left_uint64
[uint64]$val = [uint64]"0x8000000000000000"
$res = [System.Numerics.BitOperations]::RotateLeft($val, 1)
if ($res -ne [uint64]1) {
    Write-Host "FAIL: Rotate left uint64 failed"
    exit 1
}
Write-Host "PASS"
exit 0
