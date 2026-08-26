# vybe-test: powershell/type_char_classification_methods/char_in_string_indexing
$str = "Hello"
$c = $str[1]
if ($c -ne [char]'e' -or $c.GetType().Name -ne "Char") {
    Write-Host "FAIL: String indexing character extraction failed"
    exit 1
}
Write-Host "PASS"
exit 0
