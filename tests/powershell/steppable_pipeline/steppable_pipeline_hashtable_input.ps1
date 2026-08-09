# vybe-test: powershell/steppable_pipeline/steppable_pipeline_hashtable_input
$sb = { process { $_.Key } }
$sp = $sb.GetSteppablePipeline()
$sp.Begin($true)
$res = $sp.Process(@{ Key = "MapVal" })
$sp.End()
if ($res -ne "MapVal") {
    Write-Host "FAIL: SteppablePipeline hashtable input expected MapVal, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
