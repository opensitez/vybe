# vybe-test: powershell/math_rounding_midpoint_modes/round_near_epsilon_midpoint
$val = 2.5
$resAway = [math]::Round($val, [System.MidpointRounding]::AwayFromZero)
$resEven = [math]::Round($val, [System.MidpointRounding]::ToEven)
if ($resAway -ne 3 -or $resEven -ne 2) {
    Write-Host "FAIL: Near-midpoint rounding failed, resAway=$resAway, resEven=$resEven"
    exit 1
}
Write-Host "PASS"
exit 0
