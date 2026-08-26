# vybe-test: powershell/pipeline_measure_object_statistics/measure_min_and_max_on_floating_point
$floats = @(0.001, 100.5, -45.2, 3.14)
$m = $floats | Measure-Object -Minimum -Maximum
if ($m.Minimum -ne -45.2 -or $m.Maximum -ne 100.5) {
    Write-Host "FAIL: Measure-Object float Min/Max failed"
    exit 1
}
Write-Host "PASS"
exit 0
