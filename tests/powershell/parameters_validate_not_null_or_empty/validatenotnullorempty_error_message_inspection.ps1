# vybe-test: powershell/parameters_validate_not_null_or_empty/validatenotnullorempty_error_message_inspection
function Check-ErrTarget {
    param([ValidateNotNullOrEmpty()][string]$Title)
}
$msg = ""
try {
    Check-ErrTarget -Title ""
} catch {
    $msg = $_.Exception.Message
}
if (-not ($msg.Contains("Title") -or $msg.Contains("empty") -or $msg.Contains("null"))) {
    Write-Host "FAIL: Error message should mention empty/null constraint, got '$msg'"
    exit 1
}
Write-Host "PASS"
exit 0
