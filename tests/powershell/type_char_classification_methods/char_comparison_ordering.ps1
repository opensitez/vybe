# vybe-test: powershell/type_char_classification_methods/char_comparison_ordering
$c1 = [char]'a'
$c2 = [char]'b'
if (-not ($c1 -lt $c2)) {
    Write-Host "FAIL: 'a' should be less than 'b'"
    exit 1
}
Write-Host "PASS"
exit 0
