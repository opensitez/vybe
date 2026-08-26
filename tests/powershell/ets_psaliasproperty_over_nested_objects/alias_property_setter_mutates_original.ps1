# vybe-test: powershell/ets_psaliasproperty_over_nested_objects/alias_property_setter_mutates_original
$obj = [pscustomobject]@{ Title = "Draft" }
$obj.PSObject.Properties.Add([System.Management.Automation.PSAliasProperty]::new("Headline", "Title"))
$obj.Headline = "Published"
if ($obj.Title -ne "Published" -or $obj.Headline -ne "Published") {
    Write-Host "FAIL: PSAliasProperty setter mutation failed"
    exit 1
}
Write-Host "PASS"
exit 0
