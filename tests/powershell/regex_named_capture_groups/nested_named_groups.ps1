# vybe-test: powershell/regex_named_capture_groups/nested_named_groups
$str = "area: 123-456"
$matched = $str -match "area:\s*(?<full>(?<part1>\d{3})-(?<part2>\d{3}))"
if ($Matches.full -ne "123-456" -or $Matches.part1 -ne "123" -or $Matches.part2 -ne "456") {
    Write-Host "FAIL: Nested named groups match failed"
    exit 1
}
Write-Host "PASS"
exit 0
