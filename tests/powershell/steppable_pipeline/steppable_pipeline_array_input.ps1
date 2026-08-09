# vybe-test: powershell/steppable_pipeline/steppable_pipeline_array_input
$sb = { process { $_.Count } }
$sp = $sb.GetSteppablePipeline()
$sp.Begin($true)
$res = $sp.Process(@(1, 2, 3))
$sp.End()
if ($res -ne 3) {
    Write-Host "FAIL: SteppablePipeline array input Count expected 3, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
