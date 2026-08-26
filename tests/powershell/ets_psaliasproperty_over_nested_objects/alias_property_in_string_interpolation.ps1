# vybe-test: powershell/ets_psaliasproperty_over_nested_objects/alias_property_in_string_interpolation
$obj = [pscustomobject]@{ FirstName = "Alice" }
$obj.PSObject.Properties.Add([System.Management.Automation.PSAliasProperty]::new("Name", "FirstName"))
$str = "$($obj.Name)"
if ($str -ne "Alice") {
    Write-Host "FAIL: PSAliasProperty in string interpolation failed"
    exit 1
}
Write-Host "PASS"
exit 0
