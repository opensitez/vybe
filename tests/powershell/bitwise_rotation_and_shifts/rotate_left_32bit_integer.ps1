# vybe-test: powershell/bitwise_rotation_and_shifts/rotate_left_32bit_integer
[uint32]$val = 2147483649 # 0x80000001
$res = [System.Numerics.BitOperations]::RotateLeft($val, 1)
if ($res -ne [uint32]3) {
    Write-Host "FAIL: Rotate left failed"
    exit 1
}
Write-Host "PASS"
exit 0
