# vybe-test: powershell/argument_transformations/argument_transformation_multiple_params
class AddOneTransform : System.Management.Automation.ArgumentTransformationAttribute {
    [object] Transform([System.Management.Automation.EngineIntrinsics]$e, [object]$i) {
        return [int]$i + 1
    }
}
function Test-MultiTrans {
    param(
        [AddOneTransform()][int]$A,
        [AddOneTransform()][int]$B
    )
    return $A + $B
}
$res = Test-MultiTrans 10 20
if ($res -ne 32) {
    Write-Host "FAIL: multiple transformed parameters expected (11+21)=32, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
