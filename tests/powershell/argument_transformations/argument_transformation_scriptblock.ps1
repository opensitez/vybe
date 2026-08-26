# vybe-test: powershell/argument_transformations/argument_transformation_scriptblock
class ScriptBlockTransform : System.Management.Automation.ArgumentTransformationAttribute {
    [object] Transform([System.Management.Automation.EngineIntrinsics]$e, [object]$i) {
        if ($i -is [string]) { return [scriptblock]::Create($i) }
        return $i
    }
}
function Test-SbTrans {
    param([ScriptBlockTransform()][scriptblock]$Code)
    return &$Code
}
$res = Test-SbTrans "10 + 20"
if ($res -ne 30) {
    Write-Host "FAIL: ScriptBlockTransform expected 30, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
