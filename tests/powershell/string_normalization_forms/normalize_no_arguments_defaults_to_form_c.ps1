# vybe-test: powershell/string_normalization_forms/normalize_no_arguments_defaults_to_form_c
$decomposed = "a`u{0308}" # a + umlaut
$nDefault = $decomposed.Normalize()
$nFormC = $decomposed.Normalize([System.Text.NormalizationForm]::FormC)
if ($nDefault -ne $nFormC -or $nDefault -ne "`u{00E4}") {
    Write-Host "FAIL: Normalize() default FormC mismatch"
    exit 1
}
Write-Host "PASS"
exit 0
