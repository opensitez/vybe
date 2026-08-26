# vybe-test: powershell/bitwise_rotation_and_shifts/rotate_left_uint32_1_bit
[uint32]$val = 2147483648 # 0x80000000
$res = [System.Numerics.BitOperations]::RotateLeft($val, 1)
if ($res -ne [uint32]1) {
    Write-Host "FAIL: Rotate left 1 bit failed"
    exit 1
}
Write-Host "PASS"
exit 0
