# vybe-test: powershell/ets_psaliasproperty_over_nested_objects/alias_property_removal
$obj = [pscustomobject]@{ Val = 100 }
$obj.PSObject.Properties.Add([System.Management.Automation.PSAliasProperty]::new("AltVal", "Val"))
$obj.PSObject.Properties.Remove("AltVal")
if ($obj.AltVal -ne $null) {
    Write-Host "FAIL: PSAliasProperty removal failed"
    exit 1
}
Write-Host "PASS"
exit 0
