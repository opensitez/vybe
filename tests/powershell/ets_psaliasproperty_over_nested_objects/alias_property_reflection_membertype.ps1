# vybe-test: powershell/ets_psaliasproperty_over_nested_objects/alias_property_reflection_membertype
$prop = [System.Management.Automation.PSAliasProperty]::new("A", "B")
if ($prop.MemberType -ne [System.Management.Automation.PSMemberTypes]::AliasProperty) {
    Write-Host "FAIL: PSAliasProperty MemberType failed"
    exit 1
}
Write-Host "PASS"
exit 0
