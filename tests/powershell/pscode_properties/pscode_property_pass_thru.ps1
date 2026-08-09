# vybe-test: powershell/pscode_properties/pscode_property_pass_thru
class PassCodeHelper {
    static [string] GetP([object]$t) { return "PassVal" }
}
$g = [PassCodeHelper].GetMethod("GetP")
$obj = [pscustomobject]@{}
$res = $obj | Add-Member -MemberType CodeProperty -Name "P" -Value $g -PassThru
if ($res.P -ne "PassVal") {
    Write-Host "FAIL: Add-Member CodeProperty -PassThru expected P='PassVal'"
    exit 1
}
Write-Host "PASS"
exit 0
