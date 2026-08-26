# vybe-test: powershell/bitwise_rotation_and_shifts/rotate_right_32bit_integer
[uint32]$val = 0x00000001
$rotated = [uint32](($val -shr 1) -bor ($val -shl 31))
if ($rotated -ne [uint32]0x80000000) {
    Write-Host "FAIL: Rotate right failed"
    exit 1
}
Write-Host "PASS"
exit 0
