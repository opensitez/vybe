# vybe-test: powershell/ets_pscodeproperty_static_method_adapter/pscodeproperty_membertype_check
class DummyAdapter {
    static [string]GetDummy([psobject]$inst) { return "dummy" }
}
$obj = [pscustomobject]@{ Dummy = 1 }
$m = [DummyAdapter].GetMethod("GetDummy")
$prop = [System.Management.Automation.PSCodeProperty]::new("AdaptDummy", $m)
$obj.PSObject.Properties.Add($prop)
$member = $obj.PSObject.Properties["AdaptDummy"]
if ($member.MemberType -ne [System.Management.Automation.PSMemberTypes]::CodeProperty) {
    Write-Host "FAIL: PSCodeProperty MemberType check failed"
    exit 1
}
Write-Host "PASS"
exit 0
