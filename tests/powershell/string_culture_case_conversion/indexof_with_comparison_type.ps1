# vybe-test: powershell/string_culture_case_conversion/indexof_with_comparison_type
$str = "The Quick Brown Fox"
$idx = $str.IndexOf("quick", [System.StringComparison]::OrdinalIgnoreCase)
$idxExact = $str.IndexOf("quick", [System.StringComparison]::Ordinal)
if ($idx -ne 4 -or $idxExact -ne -1) {
    Write-Host "FAIL: IndexOf with StringComparison failed"
    exit 1
}
Write-Host "PASS"
exit 0
