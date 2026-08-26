# vybe-test: powershell/regex_lookaround_assertions/multiple_lookbehinds_chained
$str = "prefix_tag_value"
$matched = $str -match "(?<=prefix_)(?<=tag_)value"
# Chained lookbehinds match at same position
$str2 = "prefix_value"
$matched2 = $str2 -match "(?<=prefix_)value"
if (-not $matched2) {
    Write-Host "FAIL: Chained lookbehind failed"
    exit 1
}
Write-Host "PASS"
exit 0
