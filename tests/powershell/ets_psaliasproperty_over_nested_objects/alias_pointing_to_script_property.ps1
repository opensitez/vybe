# vybe-test: powershell/ets_psaliasproperty_over_nested_objects/alias_pointing_to_script_property
$obj = [pscustomobject]@{ Width = 10; Height = 5 }
$obj.PSObject.Properties.Add([System.Management.Automation.PSScriptProperty]::new("Area", { $this.Width * $this.Height }))
$obj.PSObject.Properties.Add([System.Management.Automation.PSAliasProperty]::new("SurfaceArea", "Area"))
if ($obj.SurfaceArea -ne 50) {
    Write-Host "FAIL: PSAliasProperty pointing to PSScriptProperty failed"
    exit 1
}
Write-Host "PASS"
exit 0
