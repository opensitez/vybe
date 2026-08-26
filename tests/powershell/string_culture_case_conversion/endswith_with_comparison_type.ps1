# vybe-test: powershell/string_culture_case_conversion/endswith_with_comparison_type
$str = "report.PDF"
$ends = $str.EndsWith(".pdf", [System.StringComparison]::OrdinalIgnoreCase)
$endsExact = $str.EndsWith(".pdf", [System.StringComparison]::Ordinal)
if (-not $ends -or $endsExact) {
    Write-Host "FAIL: EndsWith with StringComparison failed"
    exit 1
}
Write-Host "PASS"
exit 0
