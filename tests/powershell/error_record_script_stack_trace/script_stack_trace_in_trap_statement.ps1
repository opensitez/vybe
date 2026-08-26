# vybe-test: powershell/error_record_script_stack_trace/script_stack_trace_in_trap_statement
$capturedTrace = $null
function Test-TrapTrace {
    trap {
        $script:capturedTrace = $_.ScriptStackTrace
        continue
    }
    throw "TrapTraceError"
}
Test-TrapTrace
if ($capturedTrace -eq $null -or $capturedTrace.Length -eq 0) {
    Write-Host "FAIL: ScriptStackTrace inside trap statement failed"
    exit 1
}
Write-Host "PASS"
exit 0
