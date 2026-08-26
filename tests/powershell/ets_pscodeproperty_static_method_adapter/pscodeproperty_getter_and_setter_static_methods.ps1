# vybe-test: powershell/ets_pscodeproperty_static_method_adapter/pscodeproperty_getter_and_setter_static_methods
class CodePropertyPair {
    static [int]GetScaled([psobject]$instance) {
        return $instance.RawVal * 2
    }
    static [void]SetScaled([psobject]$instance, [int]$value) {
        $instance.RawVal = [int]($value / 2)
    }
}
$obj = [pscustomobject]@{ RawVal = 10 }
$getter = [CodePropertyPair].GetMethod("GetScaled")
$setter = [CodePropertyPair].GetMethod("SetScaled")
$prop = [System.Management.Automation.PSCodeProperty]::new("Scaled", $getter, $setter)
$obj.PSObject.Properties.Add($prop)
$val1 = $obj.Scaled # 20
$obj.Scaled = 50 # sets RawVal to 25
if ($val1 -ne 20 -or $obj.RawVal -ne 25 -or $obj.Scaled -ne 50) {
    Write-Host "FAIL: PSCodeProperty getter and setter failed"
    exit 1
}
Write-Host "PASS"
exit 0
