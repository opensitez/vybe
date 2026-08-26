# vybe-test: powershell/pipeline_nested_steppable_pipeline/steppable_pipeline_with_boolean_filter
$sb = {
    param([Parameter(ValueFromPipeline=$true)][int]$N)
    process { $N * 10 }
}
$res = @(1..3 | & $sb)
if ($res.Length -ne 3 -or $res[0] -ne 10 -or $res[2] -ne 30) {
    Write-Host "FAIL: Pipeline streaming execution failed"
    exit 1
}
Write-Host "PASS"
exit 0
