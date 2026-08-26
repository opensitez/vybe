# vybe-test: powershell/math_trigonometric_functions/acos_range_minus_one_to_one
$acos1 = [math]::Acos(1.0)
$acos0 = [math]::Acos(0.0)
if ($acos1 -ne 0.0 -or [math]::Abs($acos0 - ([math]::PI / 2.0)) -gt 1e-12) {
    Write-Host "FAIL: Acos calculations failed"
    exit 1
}
Write-Host "PASS"
exit 0
