# vybe-test: powershell/string_normalization_forms/normalize_form_d_canonical_decomposition
$composed = "`u{00E9}" # é
$decomposed = $composed.Normalize([System.Text.NormalizationForm]::FormD)
if ($decomposed.Length -ne 2) {
    Write-Host "FAIL: FormD decomposition to 2 chars failed, length: $($decomposed.Length)"
    exit 1
}
Write-Host "PASS"
exit 0
