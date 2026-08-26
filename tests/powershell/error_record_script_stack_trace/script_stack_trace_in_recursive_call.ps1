# vybe-test: powershell/error_record_script_stack_trace/script_stack_trace_in_recursive_call
function Recurse-To-Crash([int]$depth) {
    if ($depth -ge 3) { throw "MaxDepth" }
    Recurse-To-Crash ($depth + 1)
}
$err = $null
try {
    Recurse-To-Crash 0
} catch {
    $err = $_
}
$trace = $err.ScriptStackTrace
$count = ([regex]::Matches($trace, "Recurse-To-Crash")).Count
if ($count -lt 3) {
    Write-Host "FAIL: Recursive stack trace frames count expected >= 3, got $count"
    exit 1
}
Write-Host "PASS"
exit 0
