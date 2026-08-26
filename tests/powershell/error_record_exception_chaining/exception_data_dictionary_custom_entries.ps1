# vybe-test: powershell/error_record_exception_chaining/exception_data_dictionary_custom_entries
$ex = [System.Exception]::new("Custom error")
$ex.Data["ErrorCode"] = 404
$ex.Data["Timestamp"] = "2026-08-26"
if ($ex.Data["ErrorCode"] -ne 404 -or $ex.Data["Timestamp"] -ne "2026-08-26") {
    Write-Host "FAIL: Exception Data dictionary manipulation failed"
    exit 1
}
Write-Host "PASS"
exit 0
