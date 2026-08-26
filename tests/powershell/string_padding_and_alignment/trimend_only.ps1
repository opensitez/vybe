# vybe-test: powershell/string_padding_and_alignment/trimend_only
$str = "trailing space   "
$trimmed = $str.TrimEnd()
if ($trimmed -ne "trailing space") {
    Write-Host "FAIL: TrimEnd failed, got '$trimmed'"
    exit 1
}
Write-Host "PASS"
exit 0
