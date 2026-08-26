# vybe-test: powershell/pipeline_nested_steppable_pipeline/steppable_pipeline_empty_process_stream
$sb = {
    param([Parameter(ValueFromPipeline=$true)][int]$N)
    process { $N * 10 }
}
$sp = $sb.GetSteppablePipeline([System.Management.Automation.CommandOrigin]::Internal)
$sp.Begin($true)
$r1 = @($sp.Process(2))
$r2 = @($sp.Process(5))
$null = $sp.End()
$sp.Dispose()
if ($r1[0] -ne 20 -or $r2[0] -ne 50) {
    Write-Host "FAIL: SteppablePipeline execution failed"
    exit 1
}
Write-Host "PASS"
exit 0
