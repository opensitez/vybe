# vybe-test: powershell/pscode_properties/pscode_property_null_getter
class NullCodeHelper {
    static [object] GetNothing([object]$t) { return $null }
}
$g = [NullCodeHelper].GetMethod("GetNothing")
$obj = [pscustomobject]@{}
$obj | Add-Member -MemberType CodeProperty -Name "Nothing" -Value $g
if ($obj.Nothing -ne $null) {
    Write-Host "FAIL: CodeProperty null getter expected null"
    exit 1
}
Write-Host "PASS"
exit 0
