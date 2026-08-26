# vybe-test: powershell/type_char_classification_methods/char_equality_case_sensitive
$c1 = [char]'A'
$c2 = [char]'a'
if ($c1 -ceq $c2) {
    Write-Host "FAIL: case-sensitive comparison of 'A' and 'a' should be false"
    exit 1
}
Write-Host "PASS"
exit 0
