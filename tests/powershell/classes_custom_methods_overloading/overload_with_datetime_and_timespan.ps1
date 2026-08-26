# vybe-test: powershell/classes_custom_methods_overloading/overload_with_datetime_and_timespan
class TimeEngine {
    [datetime]Advance([datetime]$dt, [int]$hours) { return $dt.AddHours($hours) }
    [datetime]Advance([datetime]$dt, [timespan]$ts) { return $dt + $ts }
}
$te = [TimeEngine]::new()
$base = [datetime]::Parse("2026-08-26 10:00:00")
$d1 = $te.Advance($base, 2)
$d2 = $te.Advance($base, [timespan]::FromMinutes(30))
if ($d1.Hour -ne 12 -or $d2.Minute -ne 30) {
    Write-Host "FAIL: DateTime vs TimeSpan overload failed"
    exit 1
}
Write-Host "PASS"
exit 0
