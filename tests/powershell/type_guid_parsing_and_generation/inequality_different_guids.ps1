# vybe-test: powershell/type_guid_parsing_and_generation/inequality_different_guids
$g1 = [guid]::Parse("11111111-1111-1111-1111-111111111111")
$g2 = [guid]::Parse("22222222-2222-2222-2222-222222222222")
if ($g1 -eq $g2) {
    Write-Host "FAIL: different GUIDs must compare unequal"
    exit 1
}
Write-Host "PASS"
exit 0
