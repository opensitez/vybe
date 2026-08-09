# vybe-test: powershell/psalias_properties/psalias_property_in_function
function Add-AliasProp($o, [string]$alias, [string]$target) {
    $o | Add-Member -MemberType AliasProperty -Name $alias -Value $target
}
$obj = [pscustomobject]@{ Real = "FuncData" }
Add-AliasProp $obj "Synonym" "Real"
if ($obj.Synonym -ne "FuncData") {
    Write-Host "FAIL: function attached AliasProperty expected Synonym='FuncData'"
    exit 1
}
Write-Host "PASS"
exit 0
