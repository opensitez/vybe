# vybe-test: powershell/error_record_script_stack_trace/script_stack_trace_in_pipeline_execution
function Pipe-Thrower {
    process { throw "PipeTrace" }
}
$err = $null
try {
    1..3 | Pipe-Thrower
} catch {
    $err = $_
}
if (-not $err.ScriptStackTrace.Contains("Pipe-Thrower")) {
    Write-Host "FAIL: Pipeline ScriptStackTrace failed, got '$($err.ScriptStackTrace)'"
    exit 1
}
Write-Host "PASS"
exit 0
