# vybe-test: powershell/type_timespan_arithmetic/ticks_calculation
$ts = [timespan]::FromMilliseconds(1) # 1 ms = 10,000 ticks
if ($ts.Ticks -ne 10000) {
    Write-Host "FAIL: 1ms expected 10000 ticks, got $($ts.Ticks)"
    exit 1
}
Write-Host "PASS"
exit 0
