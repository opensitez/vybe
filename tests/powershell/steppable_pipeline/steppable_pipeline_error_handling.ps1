# vybe-test: powershell/steppable_pipeline/steppable_pipeline_error_handling
$sb = { process { if ($_ -eq 0) { throw "ZeroError" } else { 100 / $_ } } }
$sp = $sb.GetSteppablePipeline()
$sp.Begin($true)
try {
    $sp.Process(0)
    Write-Host "FAIL: SteppablePipeline Process expected throw on 0"
    exit 1
} catch {
    Write-Host "PASS"
    exit 0
}
