# vybe-test: powershell/string_padding_and_alignment/composite_format_alignment_positive_right_align
$formatted = "{0,10}" -f "right"
if ($formatted -ne "     right" -or $formatted.Length -ne 10) {
    Write-Host "FAIL: Composite format right align failed, got '$formatted'"
    exit 1
}
Write-Host "PASS"
exit 0
