# vybe-test: powershell/string_culture_case_conversion/string_comparer_hashcode_case_insensitive
$cmp = [System.StringComparer]::OrdinalIgnoreCase
$h1 = $cmp.GetHashCode("Hello")
$h2 = $cmp.GetHashCode("HELLO")
if ($h1 -ne $h2) {
    Write-Host "FAIL: Case-insensitive comparer hash code must be equal"
    exit 1
}
Write-Host "PASS"
exit 0
