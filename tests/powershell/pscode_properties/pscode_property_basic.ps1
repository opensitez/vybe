# vybe-test: powershell/pscode_properties/pscode_property_basic
class CodeHelper {
    static [int] GetVal([object]$target) { return 42 }
}
$obj = [pscustomobject]@{}
$getter = [CodeHelper].GetMethod("GetVal")
$cp = [System.Management.Automation.PSCodeProperty]::new("DynamicVal", $getter)
$obj.psobject.Members.Add($cp)
if ($obj.DynamicVal -ne 42) {
    Write-Host "FAIL: PSCodeProperty GetVal expected 42, got $($obj.DynamicVal)"
    exit 1
}
Write-Host "PASS"
exit 0
