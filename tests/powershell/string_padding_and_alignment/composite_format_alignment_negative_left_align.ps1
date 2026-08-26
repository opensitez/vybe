# vybe-test: powershell/string_padding_and_alignment/composite_format_alignment_negative_left_align
$formatted = "{0,-10}" -f "left"
if ($formatted -ne "left      " -or $formatted.Length -ne 10) {
    Write-Host "FAIL: Composite format left align failed, got '$formatted'"
    exit 1
}
Write-Host "PASS"
exit 0
