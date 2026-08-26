# vybe-test: powershell/math_clamp_and_range_boundaries/clamp_double_precision
$val = [math]::Clamp(12.345, 1.0, 10.0)
if ($val -ne 10.0) {
    Write-Host "FAIL: Clamp double precision failed, got $val"
    exit 1
}
Write-Host "PASS"
exit 0
