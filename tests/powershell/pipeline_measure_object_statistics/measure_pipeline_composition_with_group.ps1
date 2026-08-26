# vybe-test: powershell/pipeline_measure_object_statistics/measure_pipeline_composition_with_group
$items = @(
    [pscustomobject]@{ Dept = "IT"; Cost = 100 },
    [pscustomobject]@{ Dept = "IT"; Cost = 200 },
    [pscustomobject]@{ Dept = "HR"; Cost = 50 }
)
$deptTotals = $items | Group-Object -Property Dept | ForEach-Object {
    [pscustomobject]@{
        Dept = $_.Name
        Total = ($_.Group | Measure-Object -Property Cost -Sum).Sum
    }
}
$itTotal = ($deptTotals | Where-Object { $_.Dept -eq "IT" }).Total
$hrTotal = ($deptTotals | Where-Object { $_.Dept -eq "HR" }).Total
if ($itTotal -ne 300 -or $hrTotal -ne 50) {
    Write-Host "FAIL: Measure-Object grouped composition failed"
    exit 1
}
Write-Host "PASS"
exit 0
