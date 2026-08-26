# vybe-test: powershell/string_padding_and_alignment/trim_start_and_end_spaces
$str = "   padded string   "
$trimmed = $str.Trim()
if ($trimmed -ne "padded string") {
    Write-Host "FAIL: Trim failed, got '$trimmed'"
    exit 1
}
Write-Host "PASS"
exit 0
