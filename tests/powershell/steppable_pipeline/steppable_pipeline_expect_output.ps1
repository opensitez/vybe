# vybe-test: powershell/steppable_pipeline/steppable_pipeline_expect_output
$sb = { process { Write-Output "Out:$_" } }
$sp = $sb.GetSteppablePipeline()
$sp.Begin($true)
$res = $sp.Process("Val")
$sp.End()
if ($res -ne "Out:Val") {
    Write-Host "FAIL: SteppablePipeline Write-Output capture expected Out:Val, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
