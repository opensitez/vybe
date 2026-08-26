# vybe-test: powershell/enums_flags_attribute/enum_flags_bitwise_and_mask_extraction
[System.FlagsAttribute()]
enum SecurityRights {
    View = 1
    Edit = 2
    Delete = 4
}
$userRights = [SecurityRights]::View -bor [SecurityRights]::Edit
$hasEdit = ($userRights -band [SecurityRights]::Edit) -eq [SecurityRights]::Edit
$hasDelete = ($userRights -band [SecurityRights]::Delete) -eq [SecurityRights]::Delete
if (-not $hasEdit -or $hasDelete) {
    Write-Host "FAIL: Bitwise AND mask extraction failed"
    exit 1
}
Write-Host "PASS"
exit 0
