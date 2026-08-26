# vybe-test: powershell/type_char_classification_methods/is_letter_or_digit_test
$ch1 = [char]'A'
$ch2 = [char]'5'
$ch3 = [char]'#'
if (-not [char]::IsLetterOrDigit($ch1) -or -not [char]::IsLetterOrDigit($ch2) -or [char]::IsLetterOrDigit($ch3)) {
    Write-Host "FAIL: IsLetterOrDigit check failed"
    exit 1
}
Write-Host "PASS"
exit 0
