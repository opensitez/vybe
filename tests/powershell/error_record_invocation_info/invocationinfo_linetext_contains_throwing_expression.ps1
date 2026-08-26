# vybe-test: powershell/error_record_invocation_info/invocationinfo_linetext_contains_throwing_expression
$err = $null
try {
    throw "ErrorCheck"
} catch {
    $err = $_
}
if ($err.InvocationInfo -eq $null -or $err.InvocationInfo.ScriptLineNumber -le 0) {
    Write-Host "FAIL: InvocationInfo check failed"
    exit 1
}
Write-Host "PASS"
exit 0
