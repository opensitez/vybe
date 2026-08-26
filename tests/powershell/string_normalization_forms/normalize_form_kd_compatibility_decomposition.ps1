# vybe-test: powershell/string_normalization_forms/normalize_form_kd_compatibility_decomposition
$fraction = "`u{00BD}" # 1/2 fraction symbol
$normalized = $fraction.Normalize([System.Text.NormalizationForm]::FormKD)
if ($normalized -ne "1/2" -and -not ($normalized.Contains("1") -and $normalized.Contains("2"))) {
    Write-Host "FAIL: FormKD normalization failed"
    exit 1
}
Write-Host "PASS"
exit 0
