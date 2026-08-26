# vybe-test: powershell/math_trigonometric_functions/atan_zero_and_one
$atan0 = [math]::Atan(0.0)
$atan1 = [math]::Atan(1.0)
if ($atan0 -ne 0.0 -or [math]::Abs($atan1 - ([math]::PI / 4.0)) -gt 1e-12) {
    Write-Host "FAIL: Atan calculations failed"
    exit 1
}
Write-Host "PASS"
exit 0
