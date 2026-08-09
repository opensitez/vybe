# vybe-test: powershell/pipeline_object_capture/capture_in_return_statement
function Get-PipelineData {
    return (10..12 | Measure-Object -Sum).Sum
}
$res = Get-PipelineData
if ($res -ne 33) {
    Write-Host "FAIL: return statement with pipeline capture expected 33, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
