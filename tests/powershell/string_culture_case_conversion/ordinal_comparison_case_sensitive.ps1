# vybe-test: powershell/string_culture_case_conversion/ordinal_comparison_case_sensitive
$cmp = [System.StringComparer]::Ordinal
$res = $cmp.Compare("abc", "ABC")
if ($res -le 0) {
    Write-Host "FAIL: Ordinal compare ('abc', 'ABC') expected > 0, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
