# vybe-test: powershell/new_object_cmdlet/new_object_timespan_argumentlist_hours_minutes
# New-Object System.TimeSpan with -ArgumentList @(hours, minutes, seconds) constructs correct span
$ts = New-Object System.TimeSpan -ArgumentList 2, 30, 0

if ($ts.Hours -ne 2 -or $ts.Minutes -ne 30) {
    Write-Host "FAIL: TimeSpan mismatch, Hours=$($ts.Hours), Minutes=$($ts.Minutes)"
    exit 1
}

Write-Host "PASS"
exit 0
