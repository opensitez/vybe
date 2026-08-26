# vybe-test: powershell/type_char_classification_methods/is_surrogate_high_and_low
$high = [char]0xD800
$low = [char]0xDC00
if (-not [char]::IsHighSurrogate($high) -or -not [char]::IsLowSurrogate($low) -or -not [char]::IsSurrogatePair($high, $low)) {
    Write-Host "FAIL: Surrogate classification failed"
    exit 1
}
Write-Host "PASS"
exit 0
