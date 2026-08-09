# vybe-test: powershell/argument_transformations/argument_transformation_pipeline_input
class SquareTransform : System.Management.Automation.ArgumentTransformationAttribute {
    [object] Transform([System.Management.Automation.EngineIntrospector]$e, [object]$i) {
        return [int]$i * [int]$i
    }
}
function Test-PipeTrans {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline=$true)]
        [SquareTransform()]
        [int]$Val
    )
    process { return $Val }
}
$res = 2..4 | Test-PipeTrans
if ($res[0] -ne 4 -or $res[1] -ne 9 -or $res[2] -ne 16) {
    Write-Host "FAIL: pipeline argument transformation expected 4, 9, 16"
    exit 1
}
Write-Host "PASS"
exit 0
