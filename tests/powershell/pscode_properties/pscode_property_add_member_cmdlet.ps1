# vybe-test: powershell/pscode_properties/pscode_property_add_member_cmdlet
class CmdletCodeHelper {
    static [string] GetGreeting([object]$t) { return "HelloCode" }
}
$obj = [pscustomobject]@{}
$g = [CmdletCodeHelper].GetMethod("GetGreeting")
$obj | Add-Member -MemberType CodeProperty -Name "Greeting" -Value $g
if ($obj.Greeting -ne "HelloCode") {
    Write-Host "FAIL: Add-Member CodeProperty expected Greeting='HelloCode', got '$($obj.Greeting)'"
    exit 1
}
Write-Host "PASS"
exit 0
