# vybe-test: powershell/argument_transformations/argument_transformation_engine_context
class IdentityTransform : System.Management.Automation.ArgumentTransformationAttribute {
    [object] Transform([System.Management.Automation.EngineIntrospector]$engineData, [object]$inputData) {
        return $inputData
    }
}
function Test-Id {
    param([IdentityTransform()]$Data)
    return $Data
}
$res = Test-Id "Unchanged"
if ($res -ne "Unchanged") {
    Write-Host "FAIL: IdentityTransform expected 'Unchanged', got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
