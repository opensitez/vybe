# vybe-test: powershell/bitwise_rotation_and_shifts/round_up_to_power_of_two
$val = [System.Numerics.BitOperations]::RoundUpToPowerOf2(17)
if ($val -ne 32) {
    Write-Host "FAIL: RoundUpToPowerOf2 expected 32, got $val"
    exit 1
}
Write-Host "PASS"
exit 0
