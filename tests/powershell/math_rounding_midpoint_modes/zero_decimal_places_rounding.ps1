# vybe-test: powershell/math_rounding_midpoint_modes/zero_decimal_places_rounding
$r = [math]::Round(7.89, 0)
if ($r -ne 8.0) {
    Write-Host "FAIL: Round with 0 decimals failed, got $r"
    exit 1
}
Write-Host "PASS"
exit 0
