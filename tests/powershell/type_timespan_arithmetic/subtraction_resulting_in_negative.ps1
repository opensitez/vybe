# vybe-test: powershell/type_timespan_arithmetic/subtraction_resulting_in_negative
$ts1 = [timespan]::FromMinutes(10)
$ts2 = [timespan]::FromMinutes(25)
$ts3 = $ts1 - $ts2
if ($ts3.TotalMinutes -ne -15.0) {
    Write-Host "FAIL: expected -15.0 minutes, got $($ts3.TotalMinutes)"
    exit 1
}
Write-Host "PASS"
exit 0
