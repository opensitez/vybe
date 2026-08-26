# vybe-test: powershell/type_guid_parsing_and_generation/new_guid_uniqueness
$g1 = [guid]::NewGuid()
$g2 = [guid]::NewGuid()
if ($g1 -eq $g2 -or $g1 -eq [guid]::Empty) {
    Write-Host "FAIL: NewGuid should produce unique non-empty GUIDs"
    exit 1
}
Write-Host "PASS"
exit 0
