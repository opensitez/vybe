# vybe-test: powershell/pipeline_measure_object_statistics/measure_negative_numbers
$m = -10, -20, -30 | Measure-Object -Sum -Average -Minimum -Maximum
if ($m.Sum -ne -60 -or $m.Average -ne -20.0 -or $m.Minimum -ne -30 -or $m.Maximum -ne -10) {
    Write-Host "FAIL: Measure-Object negative numbers failed"
    exit 1
}
Write-Host "PASS"
exit 0
