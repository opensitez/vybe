# vybe-test: powershell/pscode_properties/pscode_property_subexpression
class SubCodeHelper {
    static [int] GetCount([object]$t) { return 99 }
}
$obj = [pscustomobject]@{}
$g = [SubCodeHelper].GetMethod("GetCount")
$obj | Add-Member -MemberType CodeProperty -Name "Count" -Value $g
$msg = "Total: $( $obj.Count )"
if ($msg -ne "Total: 99") {
    Write-Host "FAIL: CodeProperty in subexpression expected 'Total: 99', got '$msg'"
    exit 1
}
Write-Host "PASS"
exit 0
