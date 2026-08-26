# vybe-test: powershell/math_clamp_and_range_boundaries/clamp_int_above_maximum
$val = [math]::Clamp(150, 0, 100)
if ($val -ne 100) {
    Write-Host "FAIL: Clamp above max expected 100, got $val"
    exit 1
}
Write-Host "PASS"
exit 0
