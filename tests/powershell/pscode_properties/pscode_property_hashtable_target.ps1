# vybe-test: powershell/pscode_properties/pscode_property_hashtable_target
class HashCodeHelper {
    static [string] GetKeyCount([object]$t) { return "Keys:$($t.Count)" }
}
$h = @{ A = 1; B = 2 }
$g = [HashCodeHelper].GetMethod("GetKeyCount")
$h | Add-Member -MemberType CodeProperty -Name "Summary" -Value $g
if ($h.Summary -ne "Keys:2") {
    Write-Host "FAIL: CodeProperty on hashtable target expected Keys:2, got $($h.Summary)"
    exit 1
}
Write-Host "PASS"
exit 0
