# vybe-test: powershell/ets_psscriptproperty_getter_and_setter/psscriptproperty_returns_hashtable
$obj = [pscustomobject]@{ K = "env"; V = "prod" }
$obj.PSObject.Properties.Add([System.Management.Automation.PSScriptProperty]::new("Map", { @{ ($this.K) = $this.V } }))
if ($obj.Map["env"] -ne "prod") {
    Write-Host "FAIL: PSScriptProperty returning hashtable failed"
    exit 1
}
Write-Host "PASS"
exit 0
