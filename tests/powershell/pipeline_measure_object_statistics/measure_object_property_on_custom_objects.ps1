# vybe-test: powershell/pipeline_measure_object_statistics/measure_object_property_on_custom_objects
$items = @(
    [pscustomobject]@{ Price = 10.5 },
    [pscustomobject]@{ Price = 20.5 },
    [pscustomobject]@{ Price = 30.0 }
)
$m = $items | Measure-Object -Property Price -Sum -Average
if ($m.Sum -ne 61.0 -or [math]::Abs($m.Average - 20.33333333) -gt 1e-4) {
    Write-Host "FAIL: Measure-Object on custom object property failed, sum=$($m.Sum), avg=$($m.Average)"
    exit 1
}
Write-Host "PASS"
exit 0
