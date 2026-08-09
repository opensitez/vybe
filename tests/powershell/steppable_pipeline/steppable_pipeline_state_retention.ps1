# vybe-test: powershell/steppable_pipeline/steppable_pipeline_state_retention
$sb = {
    begin { $script:accum = 0 }
    process { $script:accum += $_ }
    end { return $script:accum }
}
$sp = $sb.GetSteppablePipeline()
$sp.Begin($true)
$sp.Process(10)
$sp.Process(20)
$total = $sp.End()
if ($total -ne 30) {
    Write-Host "FAIL: SteppablePipeline state retention expected 30, got $total"
    exit 1
}
Write-Host "PASS"
exit 0
