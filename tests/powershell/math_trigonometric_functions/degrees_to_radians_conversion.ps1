# vybe-test: powershell/math_trigonometric_functions/degrees_to_radians_conversion
$deg = 180.0
$rad = $deg * ([math]::PI / 180.0)
if ([math]::Abs($rad - [math]::PI) -gt 1e-12) {
    Write-Host "FAIL: Degree to radian conversion failed"
    exit 1
}
Write-Host "PASS"
exit 0
