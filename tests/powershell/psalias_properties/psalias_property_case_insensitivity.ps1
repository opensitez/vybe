# vybe-test: powershell/psalias_properties/psalias_property_case_insensitivity
$obj = [pscustomobject]@{ FullPath = "/var/log" }
$obj | Add-Member -MemberType AliasProperty -Name "PathAlias" -Value "FullPath"
if ($obj.pathalias -ne "/var/log") {
    Write-Host "FAIL: case-insensitive AliasProperty expected /var/log, got '$($obj.pathalias)'"
    exit 1
}
Write-Host "PASS"
exit 0
