# vybe-test: powershell/string_padding_and_alignment/trimend_character_array
$str = "filename.txt.bak"
$trimmed = $str.TrimEnd(".bak".ToCharArray())
if (-not $trimmed.StartsWith("filename")) {
    Write-Host "FAIL: TrimEnd char array failed, got '$trimmed'"
    exit 1
}
Write-Host "PASS"
exit 0
