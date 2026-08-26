# vybe-test: powershell/pscode_properties/pscode_property_type_check
class CodeHelperTest {
    static [int] GetVal([System.Management.Automation.PSObject]$target) { return 42 }
}
$obj = [pscustomobject]@{}
$getter = [CodeHelperTest].GetMethod("GetVal", [type[]]@([System.Management.Automation.PSObject]))
$cp = [System.Management.Automation.PSCodeProperty]::new("DynamicVal", $getter)
$obj.PSObject.Members.Add($cp)
if ($obj.DynamicVal -eq 42) {
    Write-Host "PASS"
    exit 0
}
Write-Host "FAIL"
exit 1
