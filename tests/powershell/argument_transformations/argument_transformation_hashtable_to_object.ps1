# vybe-test: powershell/argument_transformations/argument_transformation_hashtable_to_object
class ObjTransform : System.Management.Automation.ArgumentTransformationAttribute {
    [object] Transform([System.Management.Automation.EngineIntrinsics]$e, [object]$i) {
        if ($i -is [hashtable]) { return [pscustomobject]$i }
        return $i
    }
}
function Test-Obj {
    param([ObjTransform()][object]$Data)
    return $Data
}
$res = Test-Obj @{ K = "V" }
if (-not ($res -is [PSCustomObject]) -or $res.K -ne "V") {
    Write-Host "FAIL: ObjTransform expected PSCustomObject K=V"
    exit 1
}
Write-Host "PASS"
exit 0
