# vybe-test: powershell/type_char_classification_methods/is_separator_category
$sp = [char]' '
if (-not [char]::IsSeparator($sp)) {
    Write-Host "FAIL: IsSeparator on space failed"
    exit 1
}
Write-Host "PASS"
exit 0
