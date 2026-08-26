# vybe-test: powershell/string_normalization_forms/roundtrip_form_c_to_form_d_and_back
$orig = "`u{00C5}" # A with ring above
$d = $orig.Normalize([System.Text.NormalizationForm]::FormD)
$c = $d.Normalize([System.Text.NormalizationForm]::FormC)
if ($c -ne $orig) {
    Write-Host "FAIL: Normalization roundtrip FormC -> FormD -> FormC failed"
    exit 1
}
Write-Host "PASS"
exit 0
