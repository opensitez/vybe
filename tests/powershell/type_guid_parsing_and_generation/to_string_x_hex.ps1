# vybe-test: powershell/type_guid_parsing_and_generation/to_string_x_hex
$g = [guid]::Parse("d3b07384-d113-40a2-b94f-a496c152a3b1")
if ($g.ToString("X") -ne "{0xd3b07384,0xd113,0x40a2,{0xb9,0x4f,0xa4,0x96,0xc1,0x52,0xa3,0xb1}}") {
    Write-Host "FAIL: ToString('X') failed"
    exit 1
}
Write-Host "PASS"
exit 0
