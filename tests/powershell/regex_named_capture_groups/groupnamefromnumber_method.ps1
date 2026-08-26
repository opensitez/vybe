# vybe-test: powershell/regex_named_capture_groups/groupnamefromnumber_method
$re = [regex]::new("(?<tag>\w+)")
$name = $re.GroupNameFromNumber(1)
if ($name -ne "tag") {
    Write-Host "FAIL: GroupNameFromNumber expected 'tag', got '$name'"
    exit 1
}
Write-Host "PASS"
exit 0
