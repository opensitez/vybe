# vybe-test: powershell/math_trigonometric_functions/asin_range_minus_one_to_one
$asin0 = [math]::Asin(0.0)
$asin1 = [math]::Asin(1.0)
if ($asin0 -ne 0.0 -or [math]::Abs($asin1 - ([math]::PI / 2.0)) -gt 1e-12) {
    Write-Host "FAIL: Asin calculations failed"
    exit 1
}
Write-Host "PASS"
exit 0
