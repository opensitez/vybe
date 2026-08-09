# vybe-test: powershell/pipeline_variable/pipeline_variable_sort_object
$res = 3..1 | Sort-Object -PipelineVariable s | ForEach-Object { $s }
if ($res[0] -ne 1 -or $res[2] -ne 3) {
    Write-Host "FAIL: Sort-Object -PipelineVariable expected 1, 2, 3"
    exit 1
}
Write-Host "PASS"
exit 0
