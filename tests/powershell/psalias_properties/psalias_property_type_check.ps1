# vybe-test: powershell/psalias_properties/psalias_property_type_check
$ap = [System.Management.Automation.PSAliasProperty]::new("AliasName", "ReferencedName")
if (-not ($ap -is [System.Management.Automation.PSAliasProperty])) {
    Write-Host "FAIL: object is not [PSAliasProperty]"
    exit 1
}
Write-Host "PASS"
exit 0
