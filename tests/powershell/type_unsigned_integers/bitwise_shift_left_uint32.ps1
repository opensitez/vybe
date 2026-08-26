# vybe-test: powershell/type_unsigned_integers/bitwise_shift_left_uint32
[uint32]$x = 1
$shifted = $x -shl 16
if ($shifted -ne 65536) {
    Write-Host "FAIL: uint32 left shift mismatch"
    exit 1
}
Write-Host "PASS"
exit 0
