# vybe-test: powershell/pipeline_sort_object_properties/sort_by_dynamic_property_name_variable
$propName = "Priority"
$items = @(
    [pscustomobject]@{ Priority = 3; Name = "C" },
    [pscustomobject]@{ Priority = 1; Name = "A" },
    [pscustomobject]@{ Priority = 2; Name = "B" }
)
$sorted = @($items | Sort-Object -Property $propName)
if ($sorted[0].Name -ne "A" -or $sorted[1].Name -ne "B" -or $sorted[2].Name -ne "C") {
    Write-Host "FAIL: Sort-Object dynamic property name failed"
    exit 1
}
Write-Host "PASS"
exit 0
