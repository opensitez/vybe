# vybe-test: powershell/pipeline_group_object_operations/group_object_case_sensitivity_flag
$items = @("apple", "Apple", "APPLE", "banana")
$groups = @($items | Group-Object -CaseSensitive)
if ($groups.Count -ne 4) {
    Write-Host "FAIL: Group-Object -CaseSensitive should produce 4 distinct groups, got $($groups.Count)"
    exit 1
}
Write-Host "PASS"
exit 0
