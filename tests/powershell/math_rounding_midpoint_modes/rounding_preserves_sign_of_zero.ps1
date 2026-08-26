# vybe-test: powershell/math_rounding_midpoint_modes/rounding_preserves_sign_of_zero
$r = [math]::Round(-0.0)
if ($r -ne 0.0) {
    Write-Host "FAIL: Rounding -0.0 failed"
    exit 1
}
Write-Host "PASS"
exit 0
