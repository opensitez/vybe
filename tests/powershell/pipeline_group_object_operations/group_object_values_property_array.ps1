# vybe-test: powershell/pipeline_group_object_operations/group_object_values_property_array
$items = @([pscustomobject]@{ Category = "Book"; Price = 15 })
$groups = @($items | Group-Object -Property Category)
if ($groups[0].Values[0] -ne "Book") {
    Write-Host "FAIL: Group-Object Values property failed"
    exit 1
}
Write-Host "PASS"
exit 0
