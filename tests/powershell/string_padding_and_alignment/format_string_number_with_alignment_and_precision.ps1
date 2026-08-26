# vybe-test: powershell/string_padding_and_alignment/format_string_number_with_alignment_and_precision
$formatted = "{0,8:N2}" -f 12.3
if ($formatted -ne "   12.30" -or $formatted.Length -ne 8) {
    Write-Host "FAIL: Number format with alignment and precision failed, got '$formatted'"
    exit 1
}
Write-Host "PASS"
exit 0
