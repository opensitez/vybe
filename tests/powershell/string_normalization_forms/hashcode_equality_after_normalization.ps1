# vybe-test: powershell/string_normalization_forms/hashcode_equality_after_normalization
$s1 = "na`u{00EF}ve"
$s2 = "nai`u{0308}ve"
$h1 = $s1.Normalize().GetHashCode()
$h2 = $s2.Normalize().GetHashCode()
if ($h1 -ne $h2) {
    Write-Host "FAIL: Normalized strings must have identical hash codes"
    exit 1
}
Write-Host "PASS"
exit 0
