# vybe-test: powershell/type_char_classification_methods/is_digit_positive_and_negative
$d = [char]'7'
$a = [char]'x'
if (-not [char]::IsDigit($d) -or [char]::IsDigit($a)) {
    Write-Host "FAIL: IsDigit check failed"
    exit 1
}
Write-Host "PASS"
exit 0
