# vybe-test: powershell/ets_psaliasproperty_over_nested_objects/alias_property_chain_alias_to_alias
$obj = [pscustomobject]@{ Original = "Target" }
$obj.PSObject.Properties.Add([System.Management.Automation.PSAliasProperty]::new("Alias1", "Original"))
$obj.PSObject.Properties.Add([System.Management.Automation.PSAliasProperty]::new("Alias2", "Alias1"))
if ($obj.Alias2 -ne "Target") {
    Write-Host "FAIL: Chained PSAliasProperty to PSAliasProperty failed"
    exit 1
}
Write-Host "PASS"
exit 0
