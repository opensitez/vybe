# vybe-test: powershell/argument_transformations/argument_transformation_custom_class
class MultiplyByTen : System.Management.Automation.ArgumentTransformationAttribute {
    [object] Transform([System.Management.Automation.EngineIntrinsics]$engine, [object]$input) {
        return [int]$input * 10
    }
}
function Test-Mult {
    param([MultiplyByTen()][int]$Num)
    return $Num
}
$res = Test-Mult 5
if ($res -ne 50) {
    Write-Host "FAIL: MultiplyByTen transformation expected 50, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
