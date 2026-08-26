# vybe-test: powershell/error_record_exception_chaining/aggregateexception_handle_predicate
$e1 = [System.Exception]::new("E1")
$e2 = [System.Exception]::new("E2")
$agg = [System.AggregateException]::new([System.Exception[]]@($e1, $e2))
if ($agg.InnerExceptions.Count -ne 2) {
    Write-Host "FAIL: AggregateException check failed"
    exit 1
}
Write-Host "PASS"
exit 0
