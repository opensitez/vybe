# vybe-test: powershell/pipeline_measure_object_statistics/measure_boolean_converted_to_int
$items = @(
    [pscustomobject]@{ Val = 1 },
    [pscustomobject]@{ Val = 0 },
    [pscustomobject]@{ Val = 1 }
)
$m = $items | Measure-Object -Property Val -Sum
if ($m.Sum -ne 2) {
    Write-Host "FAIL: Measure-Object boolean converted to int failed"
    exit 1
}
Write-Host "PASS"
exit 0
