# vybe-test: powershell/argument_transformations/argument_transformation_throw_on_invalid
class PositiveOnlyTransform : System.Management.Automation.ArgumentTransformationAttribute {
    [object] Transform([System.Management.Automation.EngineIntrospector]$e, [object]$i) {
        $val = [int]$i
        if ($val -le 0) { throw "Must be positive" }
        return $val
    }
}
function Test-Pos {
    param([PositiveOnlyTransform()][int]$N)
    return $N
}
try {
    Test-Pos -1
    Write-Host "FAIL: transformation expected throw on negative int"
    exit 1
} catch {
    Write-Host "PASS"
    exit 0
}
