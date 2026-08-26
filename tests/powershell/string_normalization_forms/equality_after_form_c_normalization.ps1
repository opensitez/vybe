# vybe-test: powershell/string_normalization_forms/equality_after_form_c_normalization
$s1 = "caf`u{00E9}" # cafe with composed e-acute
$s2 = "cafe`u{0301}" # cafe with decomposed e + combining accent
$n1 = $s1.Normalize()
$n2 = $s2.Normalize()
if ($n1 -ne $n2) {
    Write-Host "FAIL: Strings normalized to FormC must be equal"
    exit 1
}
Write-Host "PASS"
exit 0
