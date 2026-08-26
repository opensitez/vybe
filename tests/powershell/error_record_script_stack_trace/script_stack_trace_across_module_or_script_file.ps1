# vybe-test: powershell/error_record_script_stack_trace/script_stack_trace_across_module_or_script_file
$err = $null
try {
    . {
        function Inner-Dot { throw "DotTrace" }
        Inner-Dot
    }
} catch {
    $err = $_
}
if (-not $err.ScriptStackTrace.Contains("Inner-Dot")) {
    Write-Host "FAIL: Dot-sourced script stack trace failed, got '$($err.ScriptStackTrace)'"
    exit 1
}
Write-Host "PASS"
exit 0
