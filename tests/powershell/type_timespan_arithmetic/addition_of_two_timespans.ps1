# vybe-test: powershell/type_timespan_arithmetic/addition_of_two_timespans
$ts1 = [timespan]::FromHours(1.5)
$ts2 = [timespan]::FromHours(2.5)
$ts3 = $ts1 + $ts2
if ($ts3.TotalHours -ne 4.0) {
    Write-Host "FAIL: expected 4.0 hours, got $($ts3.TotalHours)"
    exit 1
}
Write-Host "PASS"
exit 0
