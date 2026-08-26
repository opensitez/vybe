# vybe-test: powershell/regex_lookaround_assertions/lookahead_with_or_alternation
$str = "test.jpg and test.png"
$re = [regex]::new("test(?=\.jpg|\.png)")
$matches = @($re.Matches($str))
if ($matches.Count -ne 2) {
    Write-Host "FAIL: Lookahead with alternation failed"
    exit 1
}
Write-Host "PASS"
exit 0
