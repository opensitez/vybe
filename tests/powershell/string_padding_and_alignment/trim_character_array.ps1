# vybe-test: powershell/string_padding_and_alignment/trim_character_array
$str = "123abc456"
[char[]]$chars = @([char]'1', [char]'2', [char]'3', [char]'4', [char]'5', [char]'6')
$trimmed = $str.Trim($chars)
if ($trimmed -ne "abc") {
    Write-Host "FAIL: Trim char array failed, got '$trimmed'"
    exit 1
}
Write-Host "PASS"
exit 0
