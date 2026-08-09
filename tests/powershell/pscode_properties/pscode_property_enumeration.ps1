# vybe-test: powershell/pscode_properties/pscode_property_enumeration
class EnumCodeHelper {
    static [int] GetE([object]$t) { return 1 }
}
$obj = [pscustomobject]@{}
$g = [EnumCodeHelper].GetMethod("GetE")
$obj | Add-Member -MemberType CodeProperty -Name "CodeP" -Value $g
$codes = $obj.psobject.Members | Where-Object { $_.MemberType -eq "CodeProperty" }
if ($codes.Count -ne 1 -or $codes[0].Name -ne "CodeP") {
    Write-Host "FAIL: CodeProperty enumeration expected CodeP"
    exit 1
}
Write-Host "PASS"
exit 0
