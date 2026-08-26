# vybe-test: powershell/string_normalization_forms/string_builder_with_normalized_text
$sb = [System.Text.StringBuilder]::new()
$null = $sb.Append("e`u{0301}".Normalize())
if ($sb.ToString() -ne "`u{00E9}") {
    Write-Host "FAIL: StringBuilder with normalized string failed"
    exit 1
}
Write-Host "PASS"
exit 0
