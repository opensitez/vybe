# vybe-test: powershell/ets_psaliasproperty_over_nested_objects/alias_property_to_non_existent_property_throws
$obj = [pscustomobject]@{ A = 1 }
$obj.PSObject.Properties.Add([System.Management.Automation.PSAliasProperty]::new("BadAlias", "MissingProp"))
if ($obj.BadAlias -ne $null) {
    Write-Host "FAIL: Bad alias property should return null"
    exit 1
}
Write-Host "PASS"
exit 0
