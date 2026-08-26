# vybe-test: powershell/string_padding_and_alignment/padleft_with_custom_char
$str = "7"
$padded = $str.PadLeft(4, [char]'0')
if ($padded -ne "0007") {
    Write-Host "FAIL: PadLeft with custom char failed, got '$padded'"
    exit 1
}
Write-Host "PASS"
exit 0
