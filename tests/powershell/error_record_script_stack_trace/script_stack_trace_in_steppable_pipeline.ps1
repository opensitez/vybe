# vybe-test: powershell/error_record_script_stack_trace/script_stack_trace_in_steppable_pipeline
$sb = { process { throw "StepTrace" } }
$sp = $sb.GetSteppablePipeline()
$sp.Begin($true)
$err = $null
try {
    $sp.Process(1)
} catch {
    $err = $_
}
if ($err.ScriptStackTrace -eq $null) {
    Write-Host "FAIL: SteppablePipeline ScriptStackTrace missing"
    exit 1
}
Write-Host "PASS"
exit 0
