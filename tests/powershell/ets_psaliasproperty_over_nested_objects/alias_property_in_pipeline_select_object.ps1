# vybe-test: powershell/ets_psaliasproperty_over_nested_objects/alias_property_in_pipeline_select_object
$obj = [pscustomobject]@{ FullName = "Alice Smith" }
$obj.PSObject.Properties.Add([System.Management.Automation.PSAliasProperty]::new("Name", "FullName"))
$res = $obj | Select-Object -Property Name
if ($res.Name -ne "Alice Smith") {
    Write-Host "FAIL: Alias property in Select-Object failed"
    exit 1
}
Write-Host "PASS"
exit 0
