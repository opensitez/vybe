# vybe-test: powershell/math_rounding_midpoint_modes/math_ceiling_positive_and_negative
$c1 = [math]::Ceiling(3.2)
$c2 = [math]::Ceiling(-3.2)
if ($c1 -ne 4.0 -or $c2 -ne -3.0) {
    Write-Host "FAIL: Ceiling calculation failed, c1=$c1, c2=$c2"
    exit 1
}
Write-Host "PASS"
exit 0
