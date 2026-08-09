# vybe-test: powershell/steppable_pipeline/steppable_pipeline_process_multiple
$sb = { process { "ITEM:$_" } }
$sp = $sb.GetSteppablePipeline()
$sp.Begin($true)
$outs = @()
1..3 | ForEach-Object { $script:outs += $sp.Process($_) }
$sp.End()
if ($outs[0] -ne "ITEM:1" -or $outs[2] -ne "ITEM:3") {
    Write-Host "FAIL: SteppablePipeline multiple Process calls expected ITEM:1, ITEM:2, ITEM:3"
    exit 1
}
Write-Host "PASS"
exit 0
