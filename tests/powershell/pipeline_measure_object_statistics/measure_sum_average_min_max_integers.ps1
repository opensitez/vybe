# vybe-test: powershell/pipeline_measure_object_statistics/measure_sum_average_min_max_integers
$m = 10, 20, 30, 40, 50 | Measure-Object -Sum -Average -Minimum -Maximum
if ($m.Count -ne 5 -or $m.Sum -ne 150 -or $m.Average -ne 30.0 -or $m.Minimum -ne 10 -or $m.Maximum -ne 50) {
    Write-Host "FAIL: Measure-Object integer statistics failed"
    exit 1
}
Write-Host "PASS"
exit 0
