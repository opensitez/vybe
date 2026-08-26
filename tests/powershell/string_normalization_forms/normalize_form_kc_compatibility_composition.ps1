# vybe-test: powershell/string_normalization_forms/normalize_form_kc_compatibility_composition
$ligature = "`u{FB01}" # fi ligature
$normalized = $ligature.Normalize([System.Text.NormalizationForm]::FormKC)
if ($normalized -ne "fi") {
    Write-Host "FAIL: FormKC decomposition of fi ligature failed, got $normalized"
    exit 1
}
Write-Host "PASS"
exit 0
