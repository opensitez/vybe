# vybe-test: powershell/ets_pscodeproperty_static_method_adapter/pscodeproperty_with_guid_calculation
class GuidCalc {
    static [guid]ComputeGuid([psobject]$i) { return [guid]::Parse($i.Hex) }
}
$obj = [pscustomobject]@{ Hex = "11111111-1111-1111-1111-111111111111" }
$obj.PSObject.Properties.Add([System.Management.Automation.PSCodeProperty]::new("ParsedGuid", [GuidCalc].GetMethod("ComputeGuid")))
if ($obj.ParsedGuid -isnot [guid] -or $obj.ParsedGuid.ToString() -ne "11111111-1111-1111-1111-111111111111") {
    Write-Host "FAIL: PSCodeProperty GUID calculation failed"
    exit 1
}
Write-Host "PASS"
exit 0
