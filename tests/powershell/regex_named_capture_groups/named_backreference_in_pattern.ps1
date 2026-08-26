# vybe-test: powershell/regex_named_capture_groups/named_backreference_in_pattern
$str = "<div>content</div>"
$matched = $str -match "<(?<tag>\w+)>.*?</\k<tag>>"
if (-not $matched -or $Matches.tag -ne "div") {
    Write-Host "FAIL: Named backreference \k<tag> failed"
    exit 1
}
Write-Host "PASS"
exit 0
