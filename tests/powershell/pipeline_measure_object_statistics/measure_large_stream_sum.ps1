# vybe-test: powershell/pipeline_measure_object_statistics/measure_large_stream_sum
$m = 1..1000 | Measure-Object -Sum
if ($m.Count -ne 1000 -or $m.Sum -ne 500500) {
    Write-Host "FAIL: Measure-Object 1..1000 sum failed, got $($m.Sum)"
    exit 1
}
Write-Host "PASS"
exit 0
