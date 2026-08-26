# vybe-test: powershell/type_char_classification_methods/char_to_integer_ascii_code
$c = [char]'Z'
$code = [int]$c
if ($code -ne 90) {
    Write-Host "FAIL: ASCII code for Z expected 90, got $code"
    exit 1
}
Write-Host "PASS"
exit 0
