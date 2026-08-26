# vybe-test: powershell/error_record_exception_chaining/aggregateexception_with_single_inner_exception
$single = [System.TimeoutException]::new("Timeout")
$agg = [System.AggregateException]::new($single)
if ($agg.InnerExceptions.Count -ne 1 -or $agg.InnerException -ne $single) {
    Write-Host "FAIL: Single inner AggregateException check failed"
    exit 1
}
Write-Host "PASS"
exit 0
