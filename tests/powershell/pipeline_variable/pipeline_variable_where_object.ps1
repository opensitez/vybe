# vybe-test: powershell/pipeline_variable/pipeline_variable_where_object
$res = 1..5 | ForEach-Object -PipelineVariable num { $_ } | Where-Object { $num % 2 -eq 0 }
if ($res[0] -ne 2 -or $res[1] -ne 4) {
    Write-Host "FAIL: Where-Object with PipelineVariable expected 2, 4"
    exit 1
}
Write-Host "PASS"
exit 0
