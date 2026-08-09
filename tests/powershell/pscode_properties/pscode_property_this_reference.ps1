# vybe-test: powershell/pscode_properties/pscode_property_this_reference
class TargetHelper {
    static [int] GetDoubleVal([object]$target) {
        return $target.Val * 2
    }
}
$obj = [pscustomobject]@{ Val = 21 }
$g = [TargetHelper].GetMethod("GetDoubleVal")
$cp = [System.Management.Automation.PSCodeProperty]::new("DoubleVal", $g)
$obj.psobject.Members.Add($cp)
if ($obj.DoubleVal -ne 42) {
    Write-Host "FAIL: PSCodeProperty target inspection expected 42, got $($obj.DoubleVal)"
    exit 1
}
Write-Host "PASS"
exit 0
