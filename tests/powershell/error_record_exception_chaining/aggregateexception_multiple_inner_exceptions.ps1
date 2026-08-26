# vybe-test: powershell/error_record_exception_chaining/aggregateexception_multiple_inner_exceptions
$ex1 = [System.InvalidOperationException]::new("Error 1")
$ex2 = [System.ArgumentNullException]::new("param1", "Error 2")
$agg = [System.AggregateException]::new("Batch failed", @($ex1, $ex2))
if ($agg.InnerExceptions.Count -ne 2 -or $agg.InnerExceptions[0] -ne $ex1 -or $agg.InnerExceptions[1] -ne $ex2) {
    Write-Host "FAIL: AggregateException InnerExceptions list failed"
    exit 1
}
Write-Host "PASS"
exit 0
