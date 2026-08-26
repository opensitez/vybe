# vybe-test: powershell/ets_psaliasproperty_over_nested_objects/alias_property_case_insensitivity
$obj = [pscustomobject]@{ Tag = "Alpha" }
$obj.PSObject.Properties.Add([System.Management.Automation.PSAliasProperty]::new("MyTag", "Tag"))
if ($obj.mytag -ne "Alpha" -or $obj.MYTAG -ne "Alpha") {
    Write-Host "FAIL: PSAliasProperty case-insensitivity failed"
    exit 1
}
Write-Host "PASS"
exit 0
