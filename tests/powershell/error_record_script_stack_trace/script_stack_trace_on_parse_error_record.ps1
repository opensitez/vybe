# vybe-test: powershell/error_record_script_stack_trace/script_stack_trace_on_parse_error_record
$err = $null
try {
    Invoke-Expression "1 + (2 *"
} catch {
    $err = $_
}
if ($err.ScriptStackTrace -eq $null) {
    Write-Host "FAIL: Parse error ScriptStackTrace missing"
    exit 1
}
Write-Host "PASS"
exit 0
