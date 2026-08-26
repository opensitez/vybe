# vybe-test: powershell/exceptions_custom_net_exception_classes/custom_exception_with_datetime_timestamp
class TimestampedException : System.Exception {
    [datetime]$OccurredAt
    TimestampedException([string]$m) : base($m) {
        $this.OccurredAt = [datetime]::UtcNow
    }
}
$te = [TimestampedException]::new("Timed out")
if ($te.OccurredAt.Year -lt 2026) {
    Write-Host "FAIL: Custom exception with DateTime timestamp failed"
    exit 1
}
Write-Host "PASS"
exit 0
