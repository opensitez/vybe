# vybe-test: powershell/type_guid_parsing_and_generation/compare_to_ordering
$g1 = [guid]::Parse("00000000-0000-0000-0000-000000000001")
$g2 = [guid]::Parse("00000000-0000-0000-0000-000000000002")
if ($g1.CompareTo($g2) -ge 0) {
    Write-Host "FAIL: CompareTo ordering check failed"
    exit 1
}
Write-Host "PASS"
exit 0
