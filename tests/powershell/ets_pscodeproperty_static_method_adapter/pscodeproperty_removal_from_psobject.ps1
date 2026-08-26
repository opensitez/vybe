# vybe-test: powershell/ets_pscodeproperty_static_method_adapter/pscodeproperty_removal_from_psobject
class RemCode { static [string]GetS([psobject]$i) { return "s" } }
$obj = [pscustomobject]@{ X = 1 }
$obj.PSObject.Properties.Add([System.Management.Automation.PSCodeProperty]::new("TempCode", [RemCode].GetMethod("GetS")))
$obj.PSObject.Properties.Remove("TempCode")
if ($obj.TempCode -ne $null) {
    Write-Host "FAIL: PSCodeProperty removal failed"
    exit 1
}
Write-Host "PASS"
exit 0
