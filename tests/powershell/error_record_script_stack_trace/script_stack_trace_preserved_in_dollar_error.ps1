# vybe-test: powershell/error_record_script_stack_trace/script_stack_trace_preserved_in_dollar_error
try {
    throw "PreserveTrace"
} catch {}
$trace = $Error[0].ScriptStackTrace
if ($trace -eq $null -or $trace.Length -eq 0) {
    Write-Host "FAIL: `$Error[0].ScriptStackTrace missing"
    exit 1
}
Write-Host "PASS"
exit 0
