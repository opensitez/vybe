# vybe-test: powershell/pipeline_group_object_operations/group_object_default_case_insensitivity
$items = @("apple", "Apple", "APPLE", "banana")
$groups = @($items | Group-Object)
if ($groups.Count -ne 2) {
    Write-Host "FAIL: Group-Object default should be case-insensitive (2 groups), got $($groups.Count)"
    exit 1
}
Write-Host "PASS"
exit 0
