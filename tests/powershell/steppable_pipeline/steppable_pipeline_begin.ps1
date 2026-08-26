# vybe-test: powershell/steppable_pipeline/steppable_pipeline_begin
$sb = { ForEach-Object { $_ * 2 } }
$sp = $sb.GetSteppablePipeline([System.Management.Automation.CommandOrigin]::Internal)
$sp.Begin($true)
$r1 = @($sp.Process(5))
$r2 = @($sp.Process(10))
$null = $sp.End()
$sp.Dispose()
if ($r1[0] -eq 10 -and $r2[0] -eq 20) {
    Write-Host "PASS"
    exit 0
}
Write-Host "FAIL"
exit 1
