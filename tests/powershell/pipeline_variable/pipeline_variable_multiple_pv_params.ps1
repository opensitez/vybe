# vybe-test: powershell/pipeline_variable/pipeline_variable_multiple_pv_params
$res = 1..1 | ForEach-Object -PipelineVariable v1 { 10 } | ForEach-Object -PipelineVariable v2 { 20 } | ForEach-Object { $v1 + $v2 }
if ($res -ne 30) {
    Write-Host "FAIL: multiple PipelineVariable sums expected 30, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
