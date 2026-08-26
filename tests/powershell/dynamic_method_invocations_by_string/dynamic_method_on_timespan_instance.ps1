# vybe-test: powershell/dynamic_method_invocations_by_string/dynamic_method_on_timespan_instance
$ts = [timespan]::FromMinutes(10)
$m = "Add"
$res = $ts.$m([timespan]::FromMinutes(5))
if ($res.TotalMinutes -ne 15.0) {
    Write-Host "FAIL: Dynamic method on TimeSpan failed"
    exit 1
}
Write-Host "PASS"
exit 0
