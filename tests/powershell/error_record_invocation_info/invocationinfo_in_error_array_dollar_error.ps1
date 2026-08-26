# vybe-test: powershell/error_record_invocation_info/invocationinfo_in_error_array_dollar_error
try {
    throw "ErrorArrayCheck"
} catch {}
$lastErr = $Error[0]
if ($lastErr.InvocationInfo -eq $null -or $lastErr.InvocationInfo.Line -eq $null) {
    Write-Host "FAIL: `$Error[0] InvocationInfo inspection failed"
    exit 1
}
Write-Host "PASS"
exit 0
