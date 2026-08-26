# vybe-test: powershell/string_normalization_forms/fullwidth_latin_to_ascii_via_form_kc
$fullwidth = "`u{FF21}`u{FF22}`u{FF23}" # Fullwidth ABC
$ascii = $fullwidth.Normalize([System.Text.NormalizationForm]::FormKC)
if ($ascii -ne "ABC") {
    Write-Host "FAIL: Fullwidth to ASCII normalization failed, got $ascii"
    exit 1
}
Write-Host "PASS"
exit 0
