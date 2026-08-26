# vybe-test: powershell/math_clamp_and_range_boundaries/clamp_int_below_minimum
$val = [math]::Clamp(-5, 0, 100)
if ($val -ne 0) {
    Write-Host "FAIL: Clamp below min expected 0, got $val"
    exit 1
}
Write-Host "PASS"
exit 0
