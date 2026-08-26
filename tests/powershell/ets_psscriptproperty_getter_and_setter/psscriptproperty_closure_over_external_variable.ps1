# vybe-test: powershell/ets_psscriptproperty_getter_and_setter/psscriptproperty_closure_over_external_variable
$taxRate = 0.2
$obj = [pscustomobject]@{ BasePrice = 100 }
$obj.PSObject.Properties.Add([System.Management.Automation.PSScriptProperty]::new("Tax", { $this.BasePrice * $taxRate }.GetNewClosure()))
if ($obj.Tax -ne 20.0) {
    Write-Host "FAIL: PSScriptProperty closure failed, got $($obj.Tax)"
    exit 1
}
Write-Host "PASS"
exit 0
