# vybe-test: powershell/pipeline_group_object_operations/group_object_groupinfotype_check
$groups = @("a", "b" | Group-Object)
if ($groups[0].GetType().Name -ne "GroupInfo") {
    Write-Host "FAIL: Group-Object result type expected GroupInfo, got $($groups[0].GetType().Name)"
    exit 1
}
Write-Host "PASS"
exit 0
