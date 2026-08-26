# vybe-test: powershell/pipeline_sort_object_properties/sort_timespans_duration_ordering
$t1 = [timespan]::FromMinutes(10)
$t2 = [timespan]::FromHours(1)
$t3 = [timespan]::FromSeconds(30)
$sorted = @($t1, $t2, $t3 | Sort-Object)
if ($sorted[0].TotalSeconds -ne 30 -or $sorted[2].TotalHours -ne 1.0) {
    Write-Host "FAIL: Sort-Object TimeSpan duration ordering failed"
    exit 1
}
Write-Host "PASS"
exit 0
