# vybe-test: powershell/ets_psscriptproperty_getter_and_setter/psscriptproperty_chained_to_another_script_property
$obj = [pscustomobject]@{ N = 5 }
$obj.PSObject.Properties.Add([System.Management.Automation.PSScriptProperty]::new("Double", { $this.N * 2 }))
$obj.PSObject.Properties.Add([System.Management.Automation.PSScriptProperty]::new("Quadruple", { $this.Double * 2 }))
if ($obj.Quadruple -ne 20) {
    Write-Host "FAIL: Chained PSScriptProperty failed, got $($obj.Quadruple)"
    exit 1
}
Write-Host "PASS"
exit 0
