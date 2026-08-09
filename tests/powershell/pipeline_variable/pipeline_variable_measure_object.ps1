# vybe-test: powershell/pipeline_variable/pipeline_variable_measure_object
$res = 1..10 | Measure-Object -Sum -PipelineVariable m | ForEach-Object { $m.Sum }
if ($res -ne 55) {
    Write-Host "FAIL: Measure-Object -PipelineVariable Sum expected 55, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
