# vybe-test: powershell/string_normalization_forms/superscript_digits_to_ascii_via_form_kd
$superscript = "`u{00B2}" # superscript 2
$norm = $superscript.Normalize([System.Text.NormalizationForm]::FormKD)
if ($norm -ne "2") {
    Write-Host "FAIL: Superscript normalization failed, got $norm"
    exit 1
}
Write-Host "PASS"
exit 0
