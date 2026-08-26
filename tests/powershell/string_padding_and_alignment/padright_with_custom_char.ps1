# vybe-test: powershell/string_padding_and_alignment/padright_with_custom_char
$str = "tag"
$padded = $str.PadRight(6, [char]'-')
if ($padded -ne "tag---") {
    Write-Host "FAIL: PadRight with custom char failed, got '$padded'"
    exit 1
}
Write-Host "PASS"
exit 0
