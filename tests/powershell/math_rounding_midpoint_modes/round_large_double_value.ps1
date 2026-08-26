# vybe-test: powershell/math_rounding_midpoint_modes/round_large_double_value
$val = 123456789.5
$r = [math]::Round($val, [System.MidpointRounding]::AwayFromZero)
if ($r -ne 123456790.0) {
    Write-Host "FAIL: Large double rounding failed, got $r"
    exit 1
}
Write-Host "PASS"
exit 0
