# vybe-test: powershell/argument_transformations/argument_transformation_array_split
class CommaSplitTransform : System.Management.Automation.ArgumentTransformationAttribute {
    [object] Transform([System.Management.Automation.EngineIntrospector]$e, [object]$i) {
        if ($i -is [string]) { return $i -split "," }
        return $i
    }
}
function Test-Split {
    param([CommaSplitTransform()][string[]]$Items)
    return $Items
}
$res = Test-Split "a,b,c"
if ($res.Count -ne 3 -or $res[1] -ne "b") {
    Write-Host "FAIL: CommaSplitTransform expected @('a','b','c')"
    exit 1
}
Write-Host "PASS"
exit 0
