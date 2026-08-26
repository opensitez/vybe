# vybe-test: powershell/pipeline_group_object_operations/group_by_single_property
$items = @(
    [pscustomobject]@{ Dept = "IT"; Name = "Alice" },
    [pscustomobject]@{ Dept = "HR"; Name = "Bob" },
    [pscustomobject]@{ Dept = "IT"; Name = "Charlie" }
)
$groups = @($items | Group-Object -Property Dept)
if ($groups.Count -ne 2) {
    Write-Host "FAIL: Group-Object single property count failed, got $($groups.Count)"
    exit 1
}
$itGroup = $groups | Where-Object { $_.Name -eq "IT" }
if ($itGroup.Count -ne 2) {
    Write-Host "FAIL: IT group count failed, got $($itGroup.Count)"
    exit 1
}
Write-Host "PASS"
exit 0
