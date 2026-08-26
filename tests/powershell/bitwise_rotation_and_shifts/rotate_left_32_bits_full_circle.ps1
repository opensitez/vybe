# vybe-test: powershell/bitwise_rotation_and_shifts/rotate_left_32_bits_full_circle
[uint32]$val = 0x12345678
$rot = [System.Numerics.BitOperations]::RotateLeft($val, 32)
if ($rot -ne $val) {
    Write-Host "FAIL: RotateLeft by 32 must return original value"
    exit 1
}
Write-Host "PASS"
exit 0
