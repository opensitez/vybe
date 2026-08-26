# vybe-test: powershell/type_char_classification_methods/is_letter_ascii_and_unicode
$ch1 = [char]'G'
$ch2 = [char]'9'
if (-not [char]::IsLetter($ch1) -or [char]::IsLetter($ch2)) {
    Write-Host "FAIL: IsLetter check failed"
    exit 1
}
Write-Host "PASS"
exit 0
