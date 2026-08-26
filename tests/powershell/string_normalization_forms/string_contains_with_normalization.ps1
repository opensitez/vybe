# vybe-test: powershell/string_normalization_forms/string_contains_with_normalization
$s = "`u{0065}`u{0301}" # e + combining acute
$normalized = $s.Normalize([System.Text.NormalizationForm]::FormC)
if ($normalized.Length -ne 1 -or $normalized -ne "`u{00E9}") {
    Write-Host "FAIL: Contains after normalization failed"
    exit 1
}
Write-Host "PASS"
exit 0
