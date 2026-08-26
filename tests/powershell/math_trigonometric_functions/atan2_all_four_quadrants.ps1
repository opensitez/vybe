# vybe-test: powershell/math_trigonometric_functions/atan2_all_four_quadrants
$q1 = [math]::Atan2(1.0, 1.0)  # pi/4
$q2 = [math]::Atan2(1.0, -1.0) # 3pi/4
if ([math]::Abs($q1 - ([math]::PI / 4.0)) -gt 1e-12 -or [math]::Abs($q2 - (3.0 * [math]::PI / 4.0)) -gt 1e-12) {
    Write-Host "FAIL: Atan2 quadrant calculations failed"
    exit 1
}
Write-Host "PASS"
exit 0
