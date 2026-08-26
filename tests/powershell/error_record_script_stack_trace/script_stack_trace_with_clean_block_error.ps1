# vybe-test: powershell/error_record_script_stack_trace/script_stack_trace_with_clean_block_error
function Clean-Crash {
    [CmdletBinding()]
    param()
    clean { throw "CleanCrash" }
}
$err = $null
try {
    Clean-Crash
} catch {
    $err = $_
}
if (-not $err.ScriptStackTrace.Contains("Clean-Crash")) {
    Write-Host "FAIL: Clean block ScriptStackTrace failed"
    exit 1
}
Write-Host "PASS"
exit 0
