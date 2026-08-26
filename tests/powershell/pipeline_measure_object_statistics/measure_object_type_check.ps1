# vybe-test: powershell/pipeline_measure_object_statistics/measure_object_type_check
$m = 1..5 | Measure-Object
if ($m.GetType().Name -ne "GenericMeasureInfo" -and $m.GetType().Name -ne "MeasureInfo") {
    Write-Host "FAIL: Measure-Object result type unexpected: $($m.GetType().Name)"
    exit 1
}
Write-Host "PASS"
exit 0
