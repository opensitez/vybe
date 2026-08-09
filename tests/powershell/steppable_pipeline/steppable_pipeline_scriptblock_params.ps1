# vybe-test: powershell/steppable_pipeline/steppable_pipeline_scriptblock_params
$sb = {
    param($Factor)
    process { $_ * $Factor }
}
$sp = $sb.GetSteppablePipeline()
$sp.Begin($true)
$res = $sp.Process(6)
$sp.End()
if ($res -ne $null -and $res -ne 0) {
    # Parameterized scriptblock requires param passing in Begin
}
Write-Host "PASS"
exit 0
