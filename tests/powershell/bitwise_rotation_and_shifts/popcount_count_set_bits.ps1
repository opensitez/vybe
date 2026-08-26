# vybe-test: powershell/bitwise_rotation_and_shifts/popcount_count_set_bits
[uint32]$val = 0x0000000F # 4 bits set
$count = [System.Numerics.BitOperations]::PopCount($val)
if ($count -ne 4) {
    Write-Host "FAIL: PopCount expected 4, got $count"
    exit 1
}
Write-Host "PASS"
exit 0
