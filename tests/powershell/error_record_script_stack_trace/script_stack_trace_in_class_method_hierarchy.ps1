# vybe-test: powershell/error_record_script_stack_trace/script_stack_trace_in_class_method_hierarchy
class TraceA {
    [void]StepA() { throw "ClassTrace" }
}
class TraceB {
    [void]StepB() { [TraceA]::new().StepA() }
}
$err = $null
try {
    [TraceB]::new().StepB()
} catch {
    $err = $_
}
$trace = $err.ScriptStackTrace
if ($trace -eq $null -or $trace.Length -eq 0) {
    Write-Host "FAIL: Class method ScriptStackTrace failed"
    exit 1
}
Write-Host "PASS"
exit 0
