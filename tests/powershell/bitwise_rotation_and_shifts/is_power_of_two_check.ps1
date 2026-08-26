# vybe-test: powershell/bitwise_rotation_and_shifts/is_power_of_two_check
$t1 = [System.Numerics.BitOperations]::IsPow2(1024)
$t2 = [System.Numerics.BitOperations]::IsPow2(1023)
if (-not $t1 -or $t2) {
    Write-Host "FAIL: IsPow2 check failed"
    exit 1
}
Write-Host "PASS"
exit 0
