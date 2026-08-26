# vybe-test: powershell/type_char_classification_methods/is_punctuation_check
$dot = [char]'.'
$comma = [char]','
$a = [char]'a'
if (-not [char]::IsPunctuation($dot) -or -not [char]::IsPunctuation($comma) -or [char]::IsPunctuation($a)) {
    Write-Host "FAIL: IsPunctuation check failed"
    exit 1
}
Write-Host "PASS"
exit 0
