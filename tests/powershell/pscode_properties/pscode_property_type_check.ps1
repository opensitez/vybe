# vybe-test: powershell/pscode_properties/pscode_property_type_check
class TypeCheckHelper {
    static [int] GetNum([object]$t) { return 100 }
}
$g = [TypeCheckHelper].GetMethod("GetNum")
$cp = [System.Management.Automation.PSCodeProperty]::new("Num", $g)
if (-not ($cp -is [System.Management.Automation.PSCodeProperty])) {
    Write-Host "FAIL: object is not [PSCodeProperty]"
    exit 1
}
Write-Host "PASS"
exit 0
