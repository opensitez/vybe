# vybe-test: powershell/type_guid_parsing_and_generation/equality_two_identical_guids
$str = "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee"
$g1 = [guid]::Parse($str)
$g2 = [guid]::Parse($str)
if ($g1 -ne $g2) {
    Write-Host "FAIL: identical GUIDs must compare equal"
    exit 1
}
Write-Host "PASS"
exit 0
