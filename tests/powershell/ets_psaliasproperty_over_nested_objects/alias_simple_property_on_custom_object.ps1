# vybe-test: powershell/ets_psaliasproperty_over_nested_objects/alias_simple_property_on_custom_object
$obj = [pscustomobject]@{ OriginalName = "Server01" }
$alias = [System.Management.Automation.PSAliasProperty]::new("AliasName", "OriginalName")
$obj.PSObject.Properties.Add($alias)
if ($obj.AliasName -ne "Server01") {
    Write-Host "FAIL: Simple PSAliasProperty read failed, got '$($obj.AliasName)'"
    exit 1
}
Write-Host "PASS"
exit 0
