# vybe-test: powershell/type_char_classification_methods/is_white_space_spaces_tabs_newlines
$space = [char]' '
$tab = [char]"`t"
$nl = [char]"`n"
$letter = [char]'Z'
if (-not [char]::IsWhiteSpace($space) -or -not [char]::IsWhiteSpace($tab) -or -not [char]::IsWhiteSpace($nl) -or [char]::IsWhiteSpace($letter)) {
    Write-Host "FAIL: IsWhiteSpace check failed"
    exit 1
}
Write-Host "PASS"
exit 0
