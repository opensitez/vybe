# vybe-test: powershell/pipeline_variable/pipeline_variable_in_function_pipeline
function Invoke-PipelineTest {
    10..11 | ForEach-Object -PipelineVariable p { $_ } | ForEach-Object { $p * 2 }
}
$res = Invoke-PipelineTest
if ($res[0] -ne 20 -or $res[1] -ne 22) {
    Write-Host "FAIL: function internal pipeline variable expected 20, 22"
    exit 1
}
Write-Host "PASS"
exit 0
