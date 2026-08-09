# vybe-test: powershell/steppable_pipeline/steppable_pipeline_end
$sb = {
    process { }
    end { "FINAL_END" }
}
$sp = $sb.GetSteppablePipeline()
$sp.Begin($true)
$endRes = $sp.End()
if ($endRes -ne "FINAL_END") {
    Write-Host "FAIL: SteppablePipeline End expected 'FINAL_END', got '$endRes'"
    exit 1
}
Write-Host "PASS"
exit 0
