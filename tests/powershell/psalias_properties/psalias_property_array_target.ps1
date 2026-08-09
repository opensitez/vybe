# vybe-test: powershell/psalias_properties/psalias_property_array_target
$arr = @(1, 2, 3)
$arr | Add-Member -MemberType AliasProperty -Name "Size" -Value "Length"
if ($arr.Size -ne 3) {
    Write-Host "FAIL: AliasProperty on array pointing to Length expected Size=3, got $($arr.Size)"
    exit 1
}
Write-Host "PASS"
exit 0
