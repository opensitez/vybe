# vybe-test: powershell/regex_named_capture_groups/optional_named_group_unmatched
$str = "name: John"
$matched = $str -match "name:\s*(?<first>\w+)(\s+(?<last>\w+))?"
if (-not $matched -or $Matches.first -ne "John" -or $Matches.ContainsKey("last")) {
    Write-Host "FAIL: Unmatched optional named group should not be populated in `$Matches"
    exit 1
}
Write-Host "PASS"
exit 0
