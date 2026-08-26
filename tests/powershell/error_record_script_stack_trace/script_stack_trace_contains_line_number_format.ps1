# vybe-test: powershell/error_record_script_stack_trace/script_stack_trace_contains_line_number_format
function Line-Checker {
    throw "CheckLineNumberFormat"
}
$err = $null
try {
    Line-Checker
} catch {
    $err = $_
}
$trace = $err.ScriptStackTrace
if (-not ($trace -match "line \d+" -or $trace -match ":\s*\d+")) {
    Write-Host "FAIL: ScriptStackTrace should indicate line numbers, got '$trace'"
    exit 1
}
Write-Host "PASS"
exit 0
