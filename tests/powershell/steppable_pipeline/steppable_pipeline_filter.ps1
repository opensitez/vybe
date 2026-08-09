# vybe-test: powershell/steppable_pipeline/steppable_pipeline_filter
filter Double-Filter { $_ * 2 }
$sb = (Get-Command Double-Filter).ScriptBlock
$sp = $sb.GetSteppablePipeline()
$sp.Begin($true)
$res = $sp.Process(9)
$sp.End()
if ($res -ne 18) {
    Write-Host "FAIL: SteppablePipeline filter expected 18, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
