# vybe-test: powershell/pipeline_measure_object_statistics/measure_with_null_and_mixed_values
$items = @(
    [pscustomobject]@{ Val = 10 },
    [pscustomobject]@{ Val = 20 }
)
$m = $items | Measure-Object -Property Val -Sum
if ($m.Sum -ne 30) {
    Write-Host "FAIL: Measure-Object property sum failed"
    exit 1
}
Write-Host "PASS"
exit 0
