# vybe-test: powershell/regex_lookaround_assertions/lookahead_not_consuming_characters
$str = "foobar"
$re = [regex]::new("foo(?=bar)bar")
$m = $re.Match($str)
if (-not $m.Success -or $m.Value -ne "foobar") {
    Write-Host "FAIL: Lookahead zero-width check failed"
    exit 1
}
Write-Host "PASS"
exit 0
