# vybe-test: powershell/pscode_properties/pscode_property_array_target
class ArrayCodeHelper {
    static [int] GetFirstElement([object]$t) { return $t[0] }
}
$arr = @(10, 20, 30)
$g = [ArrayCodeHelper].GetMethod("GetFirstElement")
$arr | Add-Member -MemberType CodeProperty -Name "First" -Value $g
if ($arr.First -ne 10) {
    Write-Host "FAIL: CodeProperty on array target expected First=10, got $($arr.First)"
    exit 1
}
Write-Host "PASS"
exit 0
