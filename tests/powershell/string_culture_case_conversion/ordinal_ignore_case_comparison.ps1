# vybe-test: powershell/string_culture_case_conversion/ordinal_ignore_case_comparison
$cmp = [System.StringComparer]::OrdinalIgnoreCase
$res = $cmp.Compare("abc", "ABC")
if ($res -ne 0) {
    Write-Host "FAIL: OrdinalIgnoreCase compare expected 0, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
