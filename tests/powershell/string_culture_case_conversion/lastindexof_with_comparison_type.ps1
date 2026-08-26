# vybe-test: powershell/string_culture_case_conversion/lastindexof_with_comparison_type
$str = "apple APPLE apple"
$idx = $str.LastIndexOf("APPLE", [System.StringComparison]::OrdinalIgnoreCase)
if ($idx -ne 12) {
    Write-Host "FAIL: LastIndexOf with StringComparison expected 12, got $idx"
    exit 1
}
Write-Host "PASS"
exit 0
