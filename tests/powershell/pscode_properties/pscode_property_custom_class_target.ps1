# vybe-test: powershell/pscode_properties/pscode_property_custom_class_target
class Entity {
    [int]$Id = 5
}
class EntityCodeHelper {
    static [string] GetCode([object]$t) { return "ENTITY-$($t.Id)" }
}
$e = [Entity]::new()
$g = [EntityCodeHelper].GetMethod("GetCode")
$e | Add-Member -MemberType CodeProperty -Name "Code" -Value $g
if ($e.Code -ne "ENTITY-5") {
    Write-Host "FAIL: CodeProperty on custom class target expected ENTITY-5, got $($e.Code)"
    exit 1
}
Write-Host "PASS"
exit 0
