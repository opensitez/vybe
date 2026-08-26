# vybe-test: powershell/regex_lookaround_assertions/negative_lookahead_with_quantifier
$str = "q without u vs quick"
$re = [regex]::new("q(?!u)\w*")
$m = $re.Match($str)
if (-not $m.Success -or $m.Value -ne "q") {
    Write-Host "FAIL: q not followed by u negative lookahead failed"
    exit 1
}
Write-Host "PASS"
exit 0
