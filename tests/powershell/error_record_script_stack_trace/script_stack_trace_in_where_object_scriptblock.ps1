# vybe-test: powershell/error_record_script_stack_trace/script_stack_trace_in_where_object_scriptblock
$err = $null
try {
    1..3 | Where-Object { throw "WhereCrash" }
} catch {
    $err = $_
}
if ($err.ScriptStackTrace -eq $null -or $err.ScriptStackTrace.Length -eq 0) {
    Write-Host "FAIL: Where-Object scriptblock ScriptStackTrace missing"
    exit 1
}
Write-Host "PASS"
exit 0
