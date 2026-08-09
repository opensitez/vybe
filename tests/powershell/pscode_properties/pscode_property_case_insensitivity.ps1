# vybe-test: powershell/pscode_properties/pscode_property_case_insensitivity
class CaseCodeHelper {
    static [string] GetValue([object]$t) { return "CaseInsensitiveVal" }
}
$obj = [pscustomobject]@{}
$g = [CaseCodeHelper].GetMethod("GetValue")
$obj | Add-Member -MemberType CodeProperty -Name "CamelCode" -Value $g
if ($obj.camelcode -ne "CaseInsensitiveVal") {
    Write-Host "FAIL: case-insensitive CodeProperty expected CaseInsensitiveVal"
    exit 1
}
Write-Host "PASS"
exit 0
