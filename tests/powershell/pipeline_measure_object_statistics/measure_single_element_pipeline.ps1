# vybe-test: powershell/pipeline_measure_object_statistics/measure_single_element_pipeline
$m = 42 | Measure-Object -Sum -Average -Minimum -Maximum
if ($m.Count -ne 1 -or $m.Sum -ne 42 -or $m.Average -ne 42.0 -or $m.Minimum -ne 42 -or $m.Maximum -ne 42) {
    Write-Host "FAIL: Measure-Object single element failed"
    exit 1
}
Write-Host "PASS"
exit 0
