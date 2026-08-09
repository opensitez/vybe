# vybe-test: powershell/steppable_pipeline/steppable_pipeline_type_check
$sb = { process { $_ } }
$sp = $sb.GetSteppablePipeline()
if (-not ($sp -is [System.Management.Automation.SteppablePipeline])) {
    Write-Host "FAIL: GetSteppablePipeline() object is not [SteppablePipeline]"
    exit 1
}
Write-Host "PASS"
exit 0
