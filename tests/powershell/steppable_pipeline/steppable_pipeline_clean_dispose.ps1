# vybe-test: powershell/steppable_pipeline/steppable_pipeline_clean_dispose
$sb = { process { $_ } }
$sp = $sb.GetSteppablePipeline()
$sp.Begin($true)
[void]$sp.Process("Item")
$sp.Dispose()
Write-Host "PASS"
exit 0
