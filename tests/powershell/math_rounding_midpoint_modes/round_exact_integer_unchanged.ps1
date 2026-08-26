# vybe-test: powershell/math_rounding_midpoint_modes/round_exact_integer_unchanged
$r = [math]::Round(10.0, 2)
if ($r -ne 10.0) {
    Write-Host "FAIL: Rounding exact integer failed"
    exit 1
}
Write-Host "PASS"
exit 0
