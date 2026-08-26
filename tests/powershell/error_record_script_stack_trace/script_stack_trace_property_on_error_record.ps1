# vybe-test: powershell/error_record_script_stack_trace/script_stack_trace_property_on_error_record
$err = $null
try {
    throw "DirectCrash"
} catch {
    $err = $_
}
if ($err.ScriptStackTrace -eq $null -or $err.ScriptStackTrace.Length -eq 0) {
    Write-Host "FAIL: ScriptStackTrace should be non-empty string"
    exit 1
}
Write-Host "PASS"
exit 0
