# vybe-test: powershell/bitwise_rotation_and_shifts/popcount_all_ones_32bit
$val = [uint32]::MaxValue
$count = [System.Numerics.BitOperations]::PopCount($val)
if ($count -ne 32) {
    Write-Host "FAIL: Popcount failed, got $count"
    exit 1
}
Write-Host "PASS"
exit 0
