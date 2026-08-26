# vybe-test: powershell/classes_constructor_overloading/constructor_with_datetime_parameter
class LogEntry {
    [datetime]$Timestamp
    [string]$Msg
    LogEntry([string]$m) { $this.Timestamp = [datetime]::UtcNow; $this.Msg = $m }
    LogEntry([datetime]$ts, [string]$m) { $this.Timestamp = $ts; $this.Msg = $m }
}
$dt = [datetime]::Parse("2026-08-26")
$e = [LogEntry]::new($dt, "Started")
if ($e.Timestamp.Year -ne 2026 -or $e.Msg -ne "Started") {
    Write-Host "FAIL: DateTime constructor failed"
    exit 1
}
Write-Host "PASS"
exit 0
