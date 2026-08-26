# vybe-test: powershell/ets_psaliasproperty_over_nested_objects/alias_property_membertype_check
$obj = [pscustomobject]@{ Data = 123 }
$obj.PSObject.Properties.Add([System.Management.Automation.PSAliasProperty]::new("AltData", "Data"))
$member = $obj.PSObject.Properties["AltData"]
if ($member.MemberType -ne [System.Management.Automation.PSMemberTypes]::AliasProperty) {
    Write-Host "FAIL: PSAliasProperty MemberType check failed"
    exit 1
}
Write-Host "PASS"
exit 0
