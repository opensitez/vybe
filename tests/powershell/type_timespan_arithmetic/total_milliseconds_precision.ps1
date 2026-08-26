# vybe-test: powershell/type_timespan_arithmetic/total_milliseconds_precision
$ts = [timespan]::FromSeconds(1.234)
if ([math]::Round($ts.TotalMilliseconds) -ne 1234) {
    Write-Host "FAIL: expected 1234 total ms, got $($ts.TotalMilliseconds)"
    exit 1
}
Write-Host "PASS"
exit 0
