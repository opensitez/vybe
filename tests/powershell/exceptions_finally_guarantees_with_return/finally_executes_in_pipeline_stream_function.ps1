# vybe-test: powershell/exceptions_finally_guarantees_with_return/finally_executes_in_pipeline_stream_function
$finallyCount = 0
function Test-PipeFinally {
    param([Parameter(ValueFromPipeline=$true)][int]$Val)
    process {
        try {
            $Val * 2
        } finally {
            $script:finallyCount++
        }
    }
}
$res = @(1, 2, 3 | Test-PipeFinally)
if ($res.Length -ne 3 -or $finallyCount -ne 3) {
    Write-Host "FAIL: Finally in pipeline process block failed, count=$finallyCount"
    exit 1
}
Write-Host "PASS"
exit 0
