# vybe-test: powershell/math_floating_point_epsilon/bitdecrement_previous_representable_double
$val = 1.0
$prev = [math]::BitDecrement($val)
if ($prev -ge $val -or ($val - $prev) -gt 1e-14) {
    Write-Host "FAIL: BitDecrement failed"
    exit 1
}
Write-Host "PASS"
exit 0
