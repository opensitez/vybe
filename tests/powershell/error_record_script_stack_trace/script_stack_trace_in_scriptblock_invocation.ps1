# vybe-test: powershell/error_record_script_stack_trace/script_stack_trace_in_scriptblock_invocation
$sbInner = { throw "SbTrace" }
$sbOuter = { & $sbInner }
$err = $null
try {
    & $sbOuter
} catch {
    $err = $_
}
if ($err.ScriptStackTrace -eq $null -or $err.ScriptStackTrace.Length -eq 0) {
    Write-Host "FAIL: ScriptBlock ScriptStackTrace failed"
    exit 1
}
Write-Host "PASS"
exit 0
