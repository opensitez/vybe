# vybe-test: powershell/string_culture_case_conversion/string_equals_with_stringcomparison_enum
$s1 = "Test"
$s2 = "test"
$eqOrdinal = $s1.Equals($s2, [System.StringComparison]::Ordinal)
$eqIgnore = $s1.Equals($s2, [System.StringComparison]::OrdinalIgnoreCase)
if ($eqOrdinal -ne $false -or $eqIgnore -ne $true) {
    Write-Host "FAIL: String.Equals with StringComparison enum failed"
    exit 1
}
Write-Host "PASS"
exit 0
