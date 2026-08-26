# vybe-test: powershell/regex_named_capture_groups/numbered_index_also_points_to_named_group
$m = [regex]::Match("2026-08-26", "(?<year>\d{4})-(?<month>\d{2})-(?<day>\d{2})")
if ($m.Groups[1].Value -ne "2026" -or $m.Groups["year"].Value -ne "2026") {
    Write-Host "FAIL: Numbered index for named group failed"
    exit 1
}
Write-Host "PASS"
exit 0
