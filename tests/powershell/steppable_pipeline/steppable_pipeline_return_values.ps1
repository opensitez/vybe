# vybe-test: powershell/steppable_pipeline/steppable_pipeline_return_values
$sb = { process { return $_ + 5 } }
$sp = $sb.GetSteppablePipeline()
$sp.Begin($true)
$res = $sp.Process(15)
$sp.End()
if ($res -ne 20) {
    Write-Host "FAIL: SteppablePipeline Process return value expected 20, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
