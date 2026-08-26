# vybe-test: powershell/string_normalization_forms/normalize_form_c_canonical_composition
$decomposed = "e`u{0301}" # e + combining acute accent
$composed = $decomposed.Normalize([System.Text.NormalizationForm]::FormC)
if ($composed.Length -ne 1 -or $composed -ne "`u{00E9}") {
    Write-Host "FAIL: FormC normalization to single char failed"
    exit 1
}
Write-Host "PASS"
exit 0
