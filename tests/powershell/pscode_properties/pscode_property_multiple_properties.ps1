# vybe-test: powershell/pscode_properties/pscode_property_multiple_properties
class MultiCodeHelper {
    static [string] GetP1([object]$t) { return "V1" }
    static [string] GetP2([object]$t) { return "V2" }
}
$g1 = [MultiCodeHelper].GetMethod("GetP1")
$g2 = [MultiCodeHelper].GetMethod("GetP2")
$obj = [pscustomobject]@{}
$obj | Add-Member -MemberType CodeProperty -Name "P1" -Value $g1
$obj | Add-Member -MemberType CodeProperty -Name "P2" -Value $g2
if ($obj.P1 -ne "V1" -or $obj.P2 -ne "V2") {
    Write-Host "FAIL: multiple CodeProperties expected P1=V1, P2=V2"
    exit 1
}
Write-Host "PASS"
exit 0
