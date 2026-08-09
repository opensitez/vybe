# vybe-test: powershell/steppable_pipeline/steppable_pipeline_closure
$mult = 3
$sb = { process { $_ * $mult } }.GetClosure()
$sp = $sb.GetSteppablePipeline()
$sp.Begin($true)
$res = $sp.Process(12)
$sp.End()
if ($res -ne 36) {
    Write-Host "FAIL: SteppablePipeline closure expected 36, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
