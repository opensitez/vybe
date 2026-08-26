# vybe-test: powershell/math_rounding_midpoint_modes/math_floor_positive_and_negative
$f1 = [math]::Floor(3.7)
$f2 = [math]::Floor(-3.7)
if ($f1 -ne 3.0 -or $f2 -ne -4.0) {
    Write-Host "FAIL: Floor calculation failed, f1=$f1, f2=$f2"
    exit 1
}
Write-Host "PASS"
exit 0
