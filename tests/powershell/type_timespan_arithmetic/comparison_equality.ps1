# vybe-test: powershell/type_timespan_arithmetic/comparison_equality
$ts1 = [timespan]::FromHours(2)
$ts2 = [timespan]::FromMinutes(120)
if ($ts1 -ne $ts2) {
    Write-Host "FAIL: 2 hours should equal 120 minutes"
    exit 1
}
Write-Host "PASS"
exit 0
