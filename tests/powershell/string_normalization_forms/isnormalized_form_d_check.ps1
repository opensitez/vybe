# vybe-test: powershell/string_normalization_forms/isnormalized_form_d_check
$composed = "`u{00E9}"
$isD = $composed.IsNormalized([System.Text.NormalizationForm]::FormD)
if ($isD) {
    Write-Host "FAIL: Composed char should not be normalized FormD"
    exit 1
}
Write-Host "PASS"
exit 0
