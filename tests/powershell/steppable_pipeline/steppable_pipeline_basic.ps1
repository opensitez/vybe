# vybe-test: powershell/steppable_pipeline/steppable_pipeline_basic
$sb = { process { $_ * 2 } }
$sp = $sb.GetSteppablePipeline()
$sp.Begin($true)
$r1 = $sp.Process(5)
$r2 = $sp.Process(10)
$sp.End()
if ($r1 -ne 10 -or $r2 -ne 20) {
    Write-Host "FAIL: SteppablePipeline basic Process expected 10, 20, got r1=$r1, r2=$r2"
    exit 1
}
Write-Host "PASS"
exit 0
