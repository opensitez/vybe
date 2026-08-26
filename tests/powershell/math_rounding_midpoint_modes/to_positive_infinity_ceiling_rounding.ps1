# vybe-test: powershell/math_rounding_midpoint_modes/to_positive_infinity_ceiling_rounding
$r1 = [math]::Round(2.1, [System.MidpointRounding]::ToPositiveInfinity)
$r2 = [math]::Round(-2.9, [System.MidpointRounding]::ToPositiveInfinity)
if ($r1 -ne 3.0 -or $r2 -ne -2.0) {
    Write-Host "FAIL: ToPositiveInfinity rounding failed, r1=$r1, r2=$r2"
    exit 1
}
Write-Host "PASS"
exit 0
