# vybe-test: powershell/error_record_invocation_info/invocationinfo_history_id_greater_than_or_equal_to_zero
$err = $null
try { throw "HistErr" } catch { $err = $_ }
if ($err.InvocationInfo.HistoryId -lt -1) {
    Write-Host "FAIL: HistoryId invalid: $($err.InvocationInfo.HistoryId)"
    exit 1
}
Write-Host "PASS"
exit 0
