# vybe-test: powershell/dynamic_property_lookup_by_variable/dynamic_property_read_on_timespan
$prop = "TotalMinutes"
$ts = [timespan]::FromMinutes(45)
$res = $ts.$prop
if ($res -ne 45.0) {
    Write-Host "FAIL: Dynamic property read on TimeSpan failed"
    exit 1
}
Write-Host "PASS"
exit 0
