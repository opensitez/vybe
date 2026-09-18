# vybe-test: powershell/engine_history_management/history_maximum_history_count_default_value
# The built-in preference variable $MaximumHistoryCount defaults to 4096 in PowerShell
if ($MaximumHistoryCount -ne 4096) {
    Write-Host "FAIL: expected `$MaximumHistoryCount to be 4096, got: $MaximumHistoryCount"
    exit 1
}

Write-Host "PASS"
exit 0
