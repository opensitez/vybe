# vybe-test: powershell/argument_transformations/argument_transformation_basic
class UpperTransformation : System.Management.Automation.ArgumentTransformationAttribute {
    [object] Transform([System.Management.Automation.EngineIntrinsics]$engineData, [object]$inputData) {
        if ($inputData -is [string]) {
            return $inputData.ToUpper()
        }
        return $inputData
    }
}
function Test-Transform {
    param(
        [UpperTransformation()]
        [string]$Name
    )
    return $Name
}
$res = Test-Transform "vybe"
if ($res -ne "VYBE") {
    Write-Host "FAIL: argument transformation expected 'VYBE', got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
