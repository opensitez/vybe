# vybe-test: powershell/argument_transformations/argument_transformation_trim_string
class LowerTrimTransform : System.Management.Automation.ArgumentTransformationAttribute {
    [object] Transform([System.Management.Automation.EngineIntrospector]$e, [object]$i) {
        if ($i -is [string]) { return $i.Trim().ToLower() }
        return $i
    }
}
function Test-LT {
    param([LowerTrimTransform()][string]$InputStr)
    return $InputStr
}
$res = Test-LT "  VYBE_FRAMEWORK  "
if ($res -ne "vybe_framework") {
    Write-Host "FAIL: LowerTrimTransform expected vybe_framework, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
