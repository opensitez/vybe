# vybe-test: powershell/ets_custom_member_sets/psmemberset_propertyset_referenced_properties
$propSet = [System.Management.Automation.PSPropertySet]::new("SummarySet", [string[]]@("Name", "Status"))
$props = @($propSet.ReferencedPropertyNames)
if ($props.Length -ne 2 -or $props[0] -ne "Name" -or $props[1] -ne "Status") {
    Write-Host "FAIL: PSPropertySet ReferencedPropertyNames failed"
    exit 1
}
Write-Host "PASS"
exit 0
