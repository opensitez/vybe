# vybe-test: powershell/steppable_pipeline/steppable_pipeline_multiple_process_calls
$sb = { process { $_ * $_ } }
$sp = $sb.GetSteppablePipeline()
$sp.Begin($true)
$r1 = $sp.Process(2)
$r2 = $sp.Process(3)
$r3 = $sp.Process(4)
$sp.End()
if ($r1 -ne 4 -or $r2 -ne 9 -or $r3 -ne 16) {
    Write-Host "FAIL: SteppablePipeline squares expected 4, 9, 16"
    exit 1
}
Write-Host "PASS"
exit 0
