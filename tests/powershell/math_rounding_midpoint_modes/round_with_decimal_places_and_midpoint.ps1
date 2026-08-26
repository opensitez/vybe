# vybe-test: powershell/math_rounding_midpoint_modes/round_with_decimal_places_and_midpoint
$r = [math]::Round(1.235, 2, [System.MidpointRounding]::AwayFromZero)
if ($r -ne 1.24) {
    Write-Host "FAIL: Round with decimals and AwayFromZero expected 1.24, got $r"
    exit 1
}
Write-Host "PASS"
exit 0
