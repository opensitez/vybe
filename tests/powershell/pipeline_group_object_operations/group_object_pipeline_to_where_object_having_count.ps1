# vybe-test: powershell/pipeline_group_object_operations/group_object_pipeline_to_where_object_having_count
$words = @("one", "two", "one", "three", "two", "one")
$dups = @($words | Group-Object | Where-Object { $_.Count -gt 1 } | ForEach-Object { $_.Name })
if ($dups.Length -ne 2 -or -not ($dups -contains "one") -or -not ($dups -contains "two")) {
    Write-Host "FAIL: Group-Object 'HAVING count > 1' pattern failed"
    exit 1
}
Write-Host "PASS"
exit 0
