# vybe-test: powershell/string_padding_and_alignment/trim_specific_characters
$str = "###important###"
$trimmed = $str.Trim([char]'#')
if ($trimmed -ne "important") {
    Write-Host "FAIL: Trim specific char failed, got '$trimmed'"
    exit 1
}
Write-Host "PASS"
exit 0
