# vybe-test: powershell/math_rounding_midpoint_modes/to_zero_truncation_rounding
$r1 = [math]::Round(2.9, [System.MidpointRounding]::ToZero)
$r2 = [math]::Round(-2.9, [System.MidpointRounding]::ToZero)
if ($r1 -ne 2.0 -or $r2 -ne -2.0) {
    Write-Host "FAIL: ToZero rounding failed, r1=$r1, r2=$r2"
    exit 1
}
Write-Host "PASS"
exit 0
