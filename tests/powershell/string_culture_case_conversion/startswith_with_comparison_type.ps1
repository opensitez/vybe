# vybe-test: powershell/string_culture_case_conversion/startswith_with_comparison_type
$str = "Configuration.json"
$starts = $str.StartsWith("config", [System.StringComparison]::OrdinalIgnoreCase)
$startsExact = $str.StartsWith("config", [System.StringComparison]::Ordinal)
if (-not $starts -or $startsExact) {
    Write-Host "FAIL: StartsWith with StringComparison failed"
    exit 1
}
Write-Host "PASS"
exit 0
