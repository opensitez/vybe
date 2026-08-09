# vybe-test: powershell/steppable_pipeline/steppable_pipeline_null_input
$sb = { process { if ($_ -eq $null) { "NULL_PROCESSED" } } }
$sp = $sb.GetSteppablePipeline()
$sp.Begin($true)
$res = $sp.Process($null)
$sp.End()
if ($res -ne "NULL_PROCESSED") {
    Write-Host "FAIL: SteppablePipeline null input expected NULL_PROCESSED, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
