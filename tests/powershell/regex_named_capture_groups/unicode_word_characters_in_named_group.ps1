# vybe-test: powershell/regex_named_capture_groups/unicode_word_characters_in_named_group
$str = "city: Z`u{00FC}rich"
$matched = $str -match "city:\s*(?<city>\w+)"
if (-not $matched -or $Matches.city -ne "Z`u{00FC}rich") {
    Write-Host "FAIL: Unicode in named capture group failed"
    exit 1
}
Write-Host "PASS"
exit 0
