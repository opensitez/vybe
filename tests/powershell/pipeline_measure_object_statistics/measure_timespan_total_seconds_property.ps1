# vybe-test: powershell/pipeline_measure_object_statistics/measure_timespan_total_seconds_property
$spans = @([timespan]::FromSeconds(10), [timespan]::FromSeconds(20), [timespan]::FromSeconds(30))
$m = $spans | Measure-Object -Property TotalSeconds -Sum -Average
if ($m.Sum -ne 60 -or $m.Average -ne 20.0) {
    Write-Host "FAIL: Measure-Object TimeSpan property failed"
    exit 1
}
Write-Host "PASS"
exit 0
