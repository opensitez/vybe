# vybe-test: powershell/error_record_script_stack_trace/script_stack_trace_in_property_getter_exception
$err = $null
try {
    throw "StackCheck"
} catch {
    $err = $_
}
if ($err.ScriptStackTrace -eq $null -or $err.ScriptStackTrace.Length -eq 0) {
    Write-Host "FAIL: ScriptStackTrace missing"
    exit 1
}
Write-Host "PASS"
exit 0
