# vybe-test: powershell/type_timespan_arithmetic/from_seconds_milliseconds
$ts = [timespan]::FromSeconds(90.5)
if ($ts.TotalSeconds -ne 90.5 -or $ts.Minutes -ne 1 -or $ts.Seconds -ne 30 -or $ts.Milliseconds -ne 500) {
    Write-Host "FAIL: FromSeconds calculation error"
    exit 1
}
Write-Host "PASS"
exit 0
