# vybe-test: powershell/error_record_invocation_info/invocationinfo_type_check
$err = $null
try { throw "Err" } catch { $err = $_ }
if ($err.InvocationInfo.GetType().Name -ne "InvocationInfo") {
    Write-Host "FAIL: InvocationInfo type check failed, got $($err.InvocationInfo.GetType().Name)"
    exit 1
}
Write-Host "PASS"
exit 0
