# vybe-test: powershell/math_clamp_and_range_boundaries/clamp_equal_min_and_max
$val = [math]::Clamp(50, 20, 20)
if ($val -ne 20) {
    Write-Host "FAIL: Clamp with min=max expected 20, got $val"
    exit 1
}
Write-Host "PASS"
exit 0
