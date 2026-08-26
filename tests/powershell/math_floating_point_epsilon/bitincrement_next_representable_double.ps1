# vybe-test: powershell/math_floating_point_epsilon/bitincrement_next_representable_double
$val = 1.0
$next = [math]::BitIncrement($val)
if ($next -le $val -or ($next - $val) -gt 1e-14) {
    Write-Host "FAIL: BitIncrement failed"
    exit 1
}
Write-Host "PASS"
exit 0
