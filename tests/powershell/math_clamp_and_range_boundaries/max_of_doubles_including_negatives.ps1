# vybe-test: powershell/math_clamp_and_range_boundaries/max_of_doubles_including_negatives
$m = [math]::Max(-15.5, -3.2)
if ($m -ne -3.2) {
    Write-Host "FAIL: Max negative doubles expected -3.2, got $m"
    exit 1
}
Write-Host "PASS"
exit 0
