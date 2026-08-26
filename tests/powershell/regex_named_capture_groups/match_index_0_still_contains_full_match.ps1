# vybe-test: powershell/regex_named_capture_groups/match_index_0_still_contains_full_match
$str = "item: 123"
$matched = $str -match "item:\s*(?<id>\d+)"
if ($Matches[0] -ne "item: 123" -or $Matches.id -ne "123") {
    Write-Host "FAIL: `$Matches[0] full match failed"
    exit 1
}
Write-Host "PASS"
exit 0
