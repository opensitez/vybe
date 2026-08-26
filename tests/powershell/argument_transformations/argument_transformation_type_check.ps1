# vybe-test: powershell/argument_transformations/argument_transformation_type_check
class EnsureIntTransform : System.Management.Automation.ArgumentTransformationAttribute {
    [object] Transform([System.Management.Automation.EngineIntrinsics]$e, [object]$i) {
        return [int]$i
    }
}
function Test-Ensure {
    param([EnsureIntTransform()]$Val)
    return $Val
}
$res = Test-Ensure "100"
if (-not ($res -is [int])) {
    Write-Host "FAIL: EnsureIntTransform output is not [int]"
    exit 1
}
Write-Host "PASS"
exit 0
