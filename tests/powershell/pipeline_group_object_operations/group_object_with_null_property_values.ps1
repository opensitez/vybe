# vybe-test: powershell/pipeline_group_object_operations/group_object_with_null_property_values
$items = @(
    [pscustomobject]@{ Tag = "A"; Id = 1 },
    [pscustomobject]@{ Tag = $null; Id = 2 },
    [pscustomobject]@{ Tag = $null; Id = 3 }
)
$groups = @($items | Group-Object -Property Tag)
if ($groups.Count -ne 2) {
    Write-Host "FAIL: Group-Object with null property failed, got $($groups.Count)"
    exit 1
}
Write-Host "PASS"
exit 0
