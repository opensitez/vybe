# vybe-test: powershell/ets_custom_member_sets/psmemberset_propertyset_membertype_check
$ps = [System.Management.Automation.PSPropertySet]::new("TestSet", [string[]]@("A"))
if ($ps.MemberType -ne [System.Management.Automation.PSMemberTypes]::PropertySet) {
    Write-Host "FAIL: PSPropertySet MemberType check failed"
    exit 1
}
Write-Host "PASS"
exit 0
