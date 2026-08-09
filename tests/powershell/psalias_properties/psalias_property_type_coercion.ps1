# vybe-test: powershell/psalias_properties/psalias_property_type_coercion
$obj = [pscustomobject]@{ StrNum = "100" }
$obj | Add-Member -MemberType AliasProperty -Name "IntNum" -Value "StrNum" -SecondValueType [int]
if ($obj.IntNum -ne 100 -or -not ($obj.IntNum -is [int])) {
    Write-Host "FAIL: AliasProperty type coercion expected int 100, got $($obj.IntNum)"
    exit 1
}
Write-Host "PASS"
exit 0
