# vybe-test: powershell/pipeline_group_object_operations/group_object_group_property_contains_actual_objects
$items = @(
    [pscustomobject]@{ Role = "Dev"; Name = "Alice" },
    [pscustomobject]@{ Role = "Dev"; Name = "Bob" }
)
$groups = @($items | Group-Object -Property Role)
$devs = $groups[0].Group
if ($devs.Count -ne 2 -or $devs[0].Name -ne "Alice" -or $devs[1].Name -ne "Bob") {
    Write-Host "FAIL: Group-Object Group property contents failed"
    exit 1
}
Write-Host "PASS"
exit 0
