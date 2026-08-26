# vybe-test: powershell/pipeline_sort_object_properties/sort_custom_objects_by_property
$items = @(
    [pscustomobject]@{ Name = "Charlie"; Age = 35 },
    [pscustomobject]@{ Name = "Alice"; Age = 25 },
    [pscustomobject]@{ Name = "Bob"; Age = 30 }
)
$sorted = @($items | Sort-Object -Property Age)
if ($sorted[0].Name -ne "Alice" -or $sorted[1].Name -ne "Bob" -or $sorted[2].Name -ne "Charlie") {
    Write-Host "FAIL: Sort-Object custom objects by property failed"
    exit 1
}
Write-Host "PASS"
exit 0
