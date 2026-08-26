# vybe-test: powershell/regex_named_capture_groups/multiple_matches_named_groups_collection
$re = [regex]::new("(?<key>\w+)=(?<val>\w+)")
$matchesColl = $re.Matches("a=1 b=2 c=3")
if ($matchesColl.Count -ne 3 -or $matchesColl[1].Groups["key"].Value -ne "b" -or $matchesColl[1].Groups["val"].Value -ne "2") {
    Write-Host "FAIL: Regex Matches collection named groups failed"
    exit 1
}
Write-Host "PASS"
exit 0
