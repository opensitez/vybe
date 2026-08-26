# vybe-test: powershell/error_record_script_stack_trace/script_stack_trace_in_catch_rethrow
function Throw-Original { throw "OriginalCrash" }
function Catch-And-Rethrow {
    try { Throw-Original } catch { throw $_ }
}
$err = $null
try {
    Catch-And-Rethrow
} catch {
    $err = $_
}
if (-not $err.ScriptStackTrace.Contains("Throw-Original")) {
    Write-Host "FAIL: Rethrown ScriptStackTrace should preserve original frame"
    exit 1
}
Write-Host "PASS"
exit 0
