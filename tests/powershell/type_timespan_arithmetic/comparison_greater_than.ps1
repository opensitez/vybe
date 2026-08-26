# vybe-test: powershell/type_timespan_arithmetic/comparison_greater_than
$ts1 = [timespan]::FromHours(3)
$ts2 = [timespan]::FromMinutes(179)
if (-not ($ts1 -gt $ts2)) {
    Write-Host "FAIL: 3 hours should be greater than 179 minutes"
    exit 1
}
Write-Host "PASS"
exit 0
