# vybe-test: powershell/ets_psaliasproperty_over_nested_objects/alias_property_with_string_interpolation
$obj = [pscustomobject]@{ Original = "TargetValue" }
$obj.PSObject.Properties.Add([System.Management.Automation.PSAliasProperty]::new("Aliased", "Original"))
$str = "Val: $($obj.Aliased)"
if ($str -ne "Val: TargetValue") {
    Write-Host "FAIL: Alias property in string interpolation failed"
    exit 1
}
Write-Host "PASS"
exit 0
