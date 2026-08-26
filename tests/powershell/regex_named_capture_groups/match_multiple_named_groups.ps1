# vybe-test: powershell/regex_named_capture_groups/match_multiple_named_groups
$str = "2026-08-26"
$matched = $str -match "(?<year>\d{4})-(?<month>\d{2})-(?<day>\d{2})"
if (-not $matched -or $Matches["year"] -ne "2026" -or $Matches["month"] -ne "08" -or $Matches["day"] -ne "26") {
    Write-Host "FAIL: Multiple named groups match failed"
    exit 1
}
Write-Host "PASS"
exit 0
