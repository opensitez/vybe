# vybe-test: powershell/regex_named_capture_groups/failed_match_does_not_mutate_existing_matches_if_checked
$str = "nomatch"
$matched = $str -match "user:(?<name>\w+)"
if ($matched) {
    Write-Host "FAIL: Non-matching string reported matched"
    exit 1
}
Write-Host "PASS"
exit 0
