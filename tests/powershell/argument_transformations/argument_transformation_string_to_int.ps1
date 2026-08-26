# vybe-test: powershell/argument_transformations/argument_transformation_string_to_int
class StringToIntTransform : System.Management.Automation.ArgumentTransformationAttribute {
    [object] Transform([System.Management.Automation.EngineIntrinsics]$e, [object]$i) {
        return [int]::Parse($i)
    }
}
function Test-Parse {
    param([StringToIntTransform()][int]$Count)
    return $Count
}
$res = Test-Parse "777"
if ($res -ne 777 -or -not ($res -is [int])) {
    Write-Host "FAIL: StringToIntTransform expected 777"
    exit 1
}
Write-Host "PASS"
exit 0
