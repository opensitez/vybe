# vybe-test: powershell/psalias_properties/psalias_property_pass_thru
$obj = [pscustomobject]@{ Actual = "PassValue" }
$res = $obj | Add-Member -MemberType AliasProperty -Name "Ref" -Value "Actual" -PassThru
if ($res.Ref -ne "PassValue") {
    Write-Host "FAIL: Add-Member AliasProperty -PassThru expected Ref='PassValue'"
    exit 1
}
Write-Host "PASS"
exit 0
