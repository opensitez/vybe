# vybe-test: powershell/string_padding_and_alignment/padleft_hex_byte_formatting
$b = 10 # 0x0A
$hex = $b.ToString("X").PadLeft(2, [char]'0')
if ($hex -ne "0A") {
    Write-Host "FAIL: PadLeft hex byte format expected 0A, got $hex"
    exit 1
}
Write-Host "PASS"
exit 0
