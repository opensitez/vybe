# vybe-test: powershell/type_timespan_arithmetic/divide_timespans
$ts1 = [timespan]::FromHours(6)
$ts2 = [timespan]::FromHours(2)
$ratio = $ts1 / $ts2
if ($ratio -ne 3.0) {
    Write-Host "FAIL: expected ratio 3.0, got $ratio"
    exit 1
}
Write-Host "PASS"
exit 0
