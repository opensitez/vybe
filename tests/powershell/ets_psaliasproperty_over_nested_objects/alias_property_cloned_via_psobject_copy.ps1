# vybe-test: powershell/ets_psaliasproperty_over_nested_objects/alias_property_cloned_via_psobject_copy
$obj = [pscustomobject]@{ HostName = "web01" }
$obj.PSObject.Properties.Add([System.Management.Automation.PSAliasProperty]::new("Server", "HostName"))
$copy = $obj.PSObject.Copy()
if ($copy.Server -ne "web01") {
    Write-Host "FAIL: PSAliasProperty on copied PSObject failed"
    exit 1
}
Write-Host "PASS"
exit 0
