# vybe-test: powershell/math_rounding_midpoint_modes/away_from_zero_rounding_negative
$r = [math]::Round(-2.5, [System.MidpointRounding]::AwayFromZero)
if ($r -ne -3.0) {
    Write-Host "FAIL: AwayFromZero negative expected -3.0, got $r"
    exit 1
}
Write-Host "PASS"
exit 0
