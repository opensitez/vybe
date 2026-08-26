# vybe-test: powershell/bitwise_rotation_and_shifts/rotate_right_uint32_1_bit
[uint32]$val = 0x00000001
# Rotate right by 1: (val >> 1) | (val << 31) => 0x80000000 = 2147483648
$rot = [System.Numerics.BitOperations]::RotateRight($val, 1)
if ($rot -ne 2147483648) {
    Write-Host "FAIL: RotateRight 1 bit failed, got $rot"
    exit 1
}
Write-Host "PASS"
exit 0
