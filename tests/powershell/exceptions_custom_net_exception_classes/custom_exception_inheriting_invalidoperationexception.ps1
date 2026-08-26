# vybe-test: powershell/exceptions_custom_net_exception_classes/custom_exception_inheriting_invalidoperationexception
class ServiceUnavailableException : System.InvalidOperationException {
    ServiceUnavailableException([string]$m) : base($m) {}
}
$sue = [ServiceUnavailableException]::new("Service 503")
if ($sue -isnot [System.InvalidOperationException] -or $sue.Message -ne "Service 503") {
    Write-Host "FAIL: Custom exception inheriting InvalidOperationException failed"
    exit 1
}
Write-Host "PASS"
exit 0
