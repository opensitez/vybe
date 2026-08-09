# vybe-test: powershell/steppable_pipeline/steppable_pipeline_begin
$sb = {
    begin { "BEGIN_OUT" }
    process { $_ }
}
$sp = $sb.GetSteppablePipeline()
$beginRes = $sp.Begin($true)
if ($beginRes -ne "BEGIN_OUT") {
    Write-Host "FAIL: SteppablePipeline Begin expected 'BEGIN_OUT', got '$beginRes'"
    exit 1
}
Write-Host "PASS"
exit 0
