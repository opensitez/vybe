# vybe-test: powershell/type_guid_parsing_and_generation/parse_hex_structure_format
$str = "{0xd3b07384,0xd113,0x40a2,{0xb9,0x4f,0xa4,0x96,0xc1,0x52,0xa3,0xb1}}"
$g = [guid]::Parse($str)
if ($g.ToString("X") -ne $str) {
    Write-Host "FAIL: hex structure parse failed"
    exit 1
}
Write-Host "PASS"
exit 0
