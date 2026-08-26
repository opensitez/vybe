# vybe-test: powershell/argument_transformations/argument_transformation_array_flattening
class FlattenTransform : System.Management.Automation.ArgumentTransformationAttribute {
    [object] Transform([System.Management.Automation.EngineIntrinsics]$e, [object]$i) {
        $flat = @()
        foreach ($elem in $i) { $flat += $elem }
        return $flat
    }
}
function Test-Flat {
    param([FlattenTransform()][object[]]$List)
    return $List.Count
}
$res = Test-Flat @(@(1, 2), @(3, 4))
if ($res -lt 2) {
    Write-Host "FAIL: FlattenTransform expected array elements count >= 2, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
