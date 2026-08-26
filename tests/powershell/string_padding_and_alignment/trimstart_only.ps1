# vybe-test: powershell/string_padding_and_alignment/trimstart_only
$str = "   leading space"
$trimmed = $str.TrimStart()
if ($trimmed -ne "leading space") {
    Write-Host "FAIL: TrimStart failed, got '$trimmed'"
    exit 1
}
Write-Host "PASS"
exit 0
