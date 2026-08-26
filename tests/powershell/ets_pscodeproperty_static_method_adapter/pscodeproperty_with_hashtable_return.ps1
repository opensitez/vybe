# vybe-test: powershell/ets_pscodeproperty_static_method_adapter/pscodeproperty_with_hashtable_return
class HtCode {
    static [hashtable]GetMap([psobject]$i) { return @{ tag = $i.Tag } }
}
$obj = [pscustomobject]@{ Tag = "v1" }
$obj.PSObject.Properties.Add([System.Management.Automation.PSCodeProperty]::new("Map", [HtCode].GetMethod("GetMap")))
if ($obj.Map["tag"] -ne "v1") {
    Write-Host "FAIL: PSCodeProperty with hashtable return failed"
    exit 1
}
Write-Host "PASS"
exit 0
