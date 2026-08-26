# vybe-test: powershell/exceptions_custom_net_exception_classes/custom_exception_with_hashtable_context
class ContextException : System.Exception {
    [hashtable]$Context
    ContextException([string]$m, [hashtable]$ctx) : base($m) {
        $this.Context = $ctx
    }
}
$ctx = @{ host = "api.prod"; retries = 3 }
$ce = [ContextException]::new("Connection dropped", $ctx)
if ($ce.Context["host"] -ne "api.prod" -or $ce.Context["retries"] -ne 3) {
    Write-Host "FAIL: Custom exception with hashtable context failed"
    exit 1
}
Write-Host "PASS"
exit 0
