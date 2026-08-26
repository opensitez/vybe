# vybe-test: powershell/ets_psaliasproperty_over_nested_objects/alias_property_is_settability_check
$obj = [pscustomobject]@{ Target = "Value123" }
$obj.PSObject.Properties.Add([System.Management.Automation.PSAliasProperty]::new("Aliased", "Target"))
if ($obj.Aliased -ne "Value123") {
    Write-Host "FAIL: Alias property failed"
    exit 1
}
Write-Host "PASS"
exit 0
