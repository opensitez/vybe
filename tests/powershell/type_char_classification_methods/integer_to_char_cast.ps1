# vybe-test: powershell/type_char_classification_methods/integer_to_char_cast
$code = 65
$ch = [char]$code
if ($ch -ne [char]'A') {
    Write-Host "FAIL: [char]65 expected A, got $ch"
    exit 1
}
Write-Host "PASS"
exit 0
