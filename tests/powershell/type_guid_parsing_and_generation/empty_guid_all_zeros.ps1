# vybe-test: powershell/type_guid_parsing_and_generation/empty_guid_all_zeros
$empty = [guid]::Empty
if ($empty.ToString() -ne "00000000-0000-0000-0000-000000000000") {
    Write-Host "FAIL: Empty GUID mismatch"
    exit 1
}
Write-Host "PASS"
exit 0
