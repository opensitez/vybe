# vybe-test: powershell/ets_psaliasproperty_over_nested_objects/alias_property_match_in_group_object
$items = @(
    [pscustomobject]@{ InternalStatus = "Active" },
    [pscustomobject]@{ InternalStatus = "Active" },
    [pscustomobject]@{ InternalStatus = "Inactive" }
)
foreach ($it in $items) {
    $it.PSObject.Properties.Add([System.Management.Automation.PSAliasProperty]::new("Status", "InternalStatus"))
}
$groups = @($items | Group-Object -Property Status)
if ($groups.Count -ne 2) {
    Write-Host "FAIL: PSAliasProperty in Group-Object failed"
    exit 1
}
Write-Host "PASS"
exit 0
