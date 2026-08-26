# vybe-test: powershell/math_rounding_midpoint_modes/to_negative_infinity_floor_rounding
$r1 = [math]::Round(2.5, [System.MidpointRounding]::ToNegativeInfinity)
$r2 = [math]::Round(-2.1, [System.MidpointRounding]::ToNegativeInfinity)
if ($r1 -ne 2.0 -or $r2 -ne -3.0) {
    Write-Host "FAIL: ToNegativeInfinity rounding failed, r1=$r1, r2=$r2"
    exit 1
}
Write-Host "PASS"
exit 0
