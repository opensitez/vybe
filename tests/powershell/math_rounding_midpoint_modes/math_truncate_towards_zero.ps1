# vybe-test: powershell/math_rounding_midpoint_modes/math_truncate_towards_zero
$t1 = [math]::Truncate(5.8)
$t2 = [math]::Truncate(-5.8)
if ($t1 -ne 5.0 -or $t2 -ne -5.0) {
    Write-Host "FAIL: Truncate failed, t1=$t1, t2=$t2"
    exit 1
}
Write-Host "PASS"
exit 0
