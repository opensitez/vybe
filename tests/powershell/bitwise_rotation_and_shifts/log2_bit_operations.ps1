# vybe-test: powershell/bitwise_rotation_and_shifts/log2_bit_operations
$l = [System.Numerics.BitOperations]::Log2(256)
if ($l -ne 8) {
    Write-Host "FAIL: BitOperations Log2 expected 8, got $l"
    exit 1
}
Write-Host "PASS"
exit 0
