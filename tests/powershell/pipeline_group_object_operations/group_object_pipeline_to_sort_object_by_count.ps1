# vybe-test: powershell/pipeline_group_object_operations/group_object_pipeline_to_sort_object_by_count
$words = @("a", "b", "c", "a", "b", "a")
$groups = @($words | Group-Object | Sort-Object -Property Count -Descending)
if ($groups[0].Name -ne "a" -or $groups[0].Count -ne 3) {
    Write-Host "FAIL: Group-Object sorted by Count descending failed"
    exit 1
}
Write-Host "PASS"
exit 0
