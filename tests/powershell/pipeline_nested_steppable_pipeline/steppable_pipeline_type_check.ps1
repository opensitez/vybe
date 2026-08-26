# vybe-test: powershell/pipeline_nested_steppable_pipeline/steppable_pipeline_type_check
$sb = { process { $_ } }
$sp = $sb.GetSteppablePipeline()
if ($sp.GetType().Name -ne "SteppablePipeline") {
    Write-Host "FAIL: GetType().Name expected SteppablePipeline, got $($sp.GetType().Name)"
    exit 1
}
Write-Host "PASS"
exit 0
