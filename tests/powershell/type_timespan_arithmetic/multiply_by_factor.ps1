# vybe-test: powershell/type_timespan_arithmetic/multiply_by_factor
$ts = [timespan]::FromMinutes(10)
$res = $ts * 3
if ($res.TotalMinutes -ne 30.0) {
    Write-Host "FAIL: expected 30 minutes, got $($res.TotalMinutes)"
    exit 1
}
Write-Host "PASS"
exit 0
