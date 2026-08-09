# vybe-test: powershell/argument_transformations/argument_transformation_null_input
class NullToZeroTransform : System.Management.Automation.ArgumentTransformationAttribute {
    [object] Transform([System.Management.Automation.EngineIntrospector]$e, [object]$i) {
        if ($i -eq $null) { return 0 }
        return $i
    }
}
function Test-NullZero {
    param([NullToZeroTransform()][int]$Value)
    return $Value
}
$res = Test-NullZero $null
if ($res -ne 0) {
    Write-Host "FAIL: NullToZeroTransform expected 0, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
