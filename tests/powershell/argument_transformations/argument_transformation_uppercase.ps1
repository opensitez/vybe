# vybe-test: powershell/argument_transformations/argument_transformation_uppercase
class TrimTransform : System.Management.Automation.ArgumentTransformationAttribute {
    [object] Transform([System.Management.Automation.EngineIntrospector]$e, [object]$i) {
        if ($i -is [string]) { return $i.Trim() }
        return $i
    }
}
function Test-Trim {
    param([TrimTransform()][string]$Text)
    return $Text
}
$res = Test-Trim "  clean  "
if ($res -ne "clean") {
    Write-Host "FAIL: TrimTransform expected 'clean', got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
