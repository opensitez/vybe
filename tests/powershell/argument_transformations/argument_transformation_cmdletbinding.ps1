# vybe-test: powershell/argument_transformations/argument_transformation_cmdletbinding
class PrefixTransform : System.Management.Automation.ArgumentTransformationAttribute {
    [object] Transform([System.Management.Automation.EngineIntrospector]$e, [object]$i) {
        return "PRE_$i"
    }
}
function Test-CmdletBindingTrans {
    [CmdletBinding()]
    param([PrefixTransform()][string]$Code)
    return $Code
}
$res = Test-CmdletBindingTrans "123"
if ($res -ne "PRE_123") {
    Write-Host "FAIL: CmdletBinding parameter transformation expected PRE_123, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
