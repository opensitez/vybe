# vybe-test: powershell/ets_psaliasproperty_over_nested_objects/alias_property_on_hashtable_wrapper
$ht = @{ RealKey = "RealValue" }
$pso = [psobject]::AsPSObject($ht)
$pso.Properties.Add([System.Management.Automation.PSAliasProperty]::new("AltKey", "RealKey"))
# Hashtable ETS wrapper property
if ($pso.Properties["AltKey"] -eq $null) {
    Write-Host "FAIL: PSAliasProperty addition to hashtable wrapper failed"
    exit 1
}
Write-Host "PASS"
exit 0
