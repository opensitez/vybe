# vybe-test: powershell/pipeline_measure_object_statistics/measure_standard_deviation
$m = 2, 4, 4, 4, 5, 5, 7, 9 | Measure-Object -StandardDeviation -Average
# StdDev of sample 2,4,4,4,5,5,7,9 = 2.138
if ($m.StandardDeviation -lt 2.0 -or $m.StandardDeviation -gt 2.3) {
    Write-Host "FAIL: StandardDeviation calculation failed, got $($m.StandardDeviation)"
    exit 1
}
Write-Host "PASS"
exit 0
