# vybe-test: powershell/math_clamp_and_range_boundaries/clamp_int_within_range
$val = [math]::Clamp(5, 1, 10)
if ($val -ne 5) {
    Write-Host "FAIL: Clamp within range failed, got $val"
    exit 1
}
Write-Host "PASS"
exit 0
