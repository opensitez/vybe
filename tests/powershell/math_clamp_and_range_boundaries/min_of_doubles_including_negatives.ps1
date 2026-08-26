# vybe-test: powershell/math_clamp_and_range_boundaries/min_of_doubles_including_negatives
$m = [math]::Min(-15.5, -3.2)
if ($m -ne -15.5) {
    Write-Host "FAIL: Min negative doubles expected -15.5, got $m"
    exit 1
}
Write-Host "PASS"
exit 0
