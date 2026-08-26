# vybe-test: powershell/type_timespan_arithmetic/negate_operator
$ts = [timespan]::FromHours(5)
$neg = $ts.Negate()
if ($neg.TotalHours -ne -5.0) {
    Write-Host "FAIL: Negate() expected -5, got $($neg.TotalHours)"
    exit 1
}
Write-Host "PASS"
exit 0
