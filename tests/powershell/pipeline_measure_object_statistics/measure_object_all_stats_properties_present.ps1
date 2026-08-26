# vybe-test: powershell/pipeline_measure_object_statistics/measure_object_all_stats_properties_present
$m = 1, 2, 3 | Measure-Object -Sum -Average -Minimum -Maximum -StandardDeviation
$props = @($m.PSObject.Properties | ForEach-Object { $_.Name })
if (-not ($props -contains "Count") -or -not ($props -contains "Sum") -or -not ($props -contains "Average") -or -not ($props -contains "StandardDeviation")) {
    Write-Host "FAIL: Measure-Object properties missing from result object"
    exit 1
}
Write-Host "PASS"
exit 0
