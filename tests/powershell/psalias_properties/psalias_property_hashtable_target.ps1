# vybe-test: powershell/psalias_properties/psalias_property_hashtable_target
$h = @{ PrimaryKey = "ID_123" }
$h | Add-Member -MemberType AliasProperty -Name "Id" -Value "PrimaryKey"
if ($h.Id -ne "ID_123") {
    Write-Host "FAIL: AliasProperty on hashtable target expected 'ID_123', got '$($h.Id)'"
    exit 1
}
Write-Host "PASS"
exit 0
