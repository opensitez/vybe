# vybe-test: powershell/pipeline_measure_object_statistics/measure_empty_pipeline_input
$arr = @()
$m = $arr | Measure-Object
if ($m.Count -ne 0) {
    Write-Host "FAIL: Measure-Object on empty pipeline failed"
    exit 1
}
Write-Host "PASS"
exit 0
