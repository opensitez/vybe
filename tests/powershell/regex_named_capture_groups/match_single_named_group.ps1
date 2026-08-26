# vybe-test: powershell/regex_named_capture_groups/match_single_named_group
$str = "user: alice"
$matched = $str -match "user:\s*(?<username>\w+)"
if (-not $matched -or $Matches["username"] -ne "alice") {
    Write-Host "FAIL: Single named group match failed"
    exit 1
}
Write-Host "PASS"
exit 0
