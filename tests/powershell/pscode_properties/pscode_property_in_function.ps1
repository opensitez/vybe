# vybe-test: powershell/pscode_properties/pscode_property_in_function
class FuncCodeHelper {
    static [string] GetStatus([object]$t) { return "FuncStatusOK" }
}
function Attach-CodeProp($o) {
    $g = [FuncCodeHelper].GetMethod("GetStatus")
    $o | Add-Member -MemberType CodeProperty -Name "CodeStatus" -Value $g
}
$obj = [pscustomobject]@{}
Attach-CodeProp $obj
if ($obj.CodeStatus -ne "FuncStatusOK") {
    Write-Host "FAIL: function attached CodeProperty expected FuncStatusOK"
    exit 1
}
Write-Host "PASS"
exit 0
