# vybe-test: powershell/error_record_exception_chaining/errorrecord_wrapping_chained_exception
$inner = [System.TimeoutException]::new("Timed out")
$outer = [System.InvalidOperationException]::new("Request failed", $inner)
$err = [System.Management.Automation.ErrorRecord]::new($outer, "ReqId", [System.Management.Automation.ErrorCategory]::OperationTimeout, $null)
if ($err.Exception.InnerException -ne $inner) {
    Write-Host "FAIL: ErrorRecord wrapping chained exception failed"
    exit 1
}
Write-Host "PASS"
exit 0
