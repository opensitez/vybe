# vybe-test: powershell/error_record_script_stack_trace/script_stack_trace_in_filter_block
filter Filter-Crash {
    throw "FilterCrash"
}
$err = $null
try {
    1 | Filter-Crash
} catch {
    $err = $_
}
if (-not $err.ScriptStackTrace.Contains("Filter-Crash")) {
    Write-Host "FAIL: Filter block ScriptStackTrace failed"
    exit 1
}
Write-Host "PASS"
exit 0
