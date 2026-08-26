# vybe-test: powershell/parameters_validate_length/validatelength_unicode_character_counting
function Set-UnicodeString {
    param([ValidateLength(1, 3)][string]$Emoji)
    return $Emoji
}
$res = Set-UnicodeString -Emoji "`u{00E9}`u{00E8}" # 2 chars
if ($res.Length -ne 2) {
    Write-Host "FAIL: Unicode string ValidateLength failed"
    exit 1
}
Write-Host "PASS"
exit 0
