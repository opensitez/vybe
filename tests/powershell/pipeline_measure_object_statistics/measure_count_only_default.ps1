# vybe-test: powershell/pipeline_measure_object_statistics/measure_count_only_default
$m = 1..10 | Measure-Object
if ($m.Count -ne 10) {
    Write-Host "FAIL: Measure-Object default count failed, got $($m.Count)"
    exit 1
}
Write-Host "PASS"
exit 0
