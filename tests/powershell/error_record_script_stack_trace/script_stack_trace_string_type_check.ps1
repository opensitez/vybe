# vybe-test: powershell/error_record_script_stack_trace/script_stack_trace_string_type_check
$err = $null
try { throw "TypeCheck" } catch { $err = $_ }
if ($err.ScriptStackTrace -isnot [string]) {
    Write-Host "FAIL: ScriptStackTrace should be System.String"
    exit 1
}
Write-Host "PASS"
exit 0
