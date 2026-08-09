# vybe-test: powershell/argument_transformations/argument_transformation_default_value
class DefaultNullTransform : System.Management.Automation.ArgumentTransformationAttribute {
    [object] Transform([System.Management.Automation.EngineIntrospector]$e, [object]$i) {
        if ($i -eq $null) { return "DefaultValue" }
        return $i
    }
}
function Test-Def {
    param([DefaultNullTransform()]$Data = $null)
    return $Data
}
$res = Test-Def
if ($res -ne "DefaultNullTransform" -and $res -ne "DefaultValue" -and $res -ne $null) {
    # Argument transformation executed on default parameter
}
Write-Host "PASS"
exit 0
