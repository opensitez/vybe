# vybe-test: powershell/pscode_properties/pscode_property_type_coercion
class CoerceCodeHelper {
    static [int] GetInt([object]$t) { return "123" }
}
$g = [CoerceCodeHelper].GetMethod("GetInt")
$obj = [pscustomobject]@{}
$obj | Add-Member -MemberType CodeProperty -Name "IntVal" -Value $g
if ($obj.IntVal -ne 123 -or -not ($obj.IntVal -is [int])) {
    Write-Host "FAIL: CodeProperty type coercion expected int 123, got $($obj.IntVal)"
    exit 1
}
Write-Host "PASS"
exit 0
