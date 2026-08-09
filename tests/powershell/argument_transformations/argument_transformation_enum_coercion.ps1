# vybe-test: powershell/argument_transformations/argument_transformation_enum_coercion
class EnumCaseTransform : System.Management.Automation.ArgumentTransformationAttribute {
    [object] Transform([System.Management.Automation.EngineIntrospector]$e, [object]$i) {
        if ($i -is [string]) { return [System.DayOfWeek]::Parse([System.DayOfWeek], $i, $true) }
        return $i
    }
}
function Test-EnumTrans {
    param([EnumCaseTransform()][System.DayOfWeek]$Day)
    return $Day
}
$res = Test-EnumTrans "friday"
if ($res -ne [System.DayOfWeek]::Friday) {
    Write-Host "FAIL: EnumCaseTransform expected Friday"
    exit 1
}
Write-Host "PASS"
exit 0
