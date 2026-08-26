# vybe-test: powershell/pipeline_nested_steppable_pipeline/steppable_pipeline_returning_array_from_single_process
$sb = {
    param([Parameter(ValueFromPipeline=$true)][int]$N)
    process { @($N, $N * 10) }
}
$sp = $sb.GetSteppablePipeline()
$sp.Begin($true)
$res = $sp.Process(5)
$sp.End()
if ($res[0] -ne 5 -or $res[1] -ne 50) {
    Write-Host "FAIL: Steppable pipeline array emission failed"
    exit 1
}
Write-Host "PASS"
exit 0
