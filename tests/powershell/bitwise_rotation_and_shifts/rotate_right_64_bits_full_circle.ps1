# vybe-test: powershell/bitwise_rotation_and_shifts/rotate_right_64_bits_full_circle
[uint64]$val = 0x123456789ABCDEF0
$rot = [System.Numerics.BitOperations]::RotateRight($val, 64)
if ($rot -ne $val) {
    Write-Host "FAIL: RotateRight by 64 must return original value"
    exit 1
}
Write-Host "PASS"
exit 0
