# vybe-test: powershell/regex_lookaround_assertions/lookahead_with_case_insensitive_flag
$re = [regex]::new("apple(?=(?i)PIE)")
$m = $re.Match("applepie")
if (-not $m.Success -or $m.Value -ne "apple") {
    Write-Host "FAIL: Case-insensitive lookahead failed"
    exit 1
}
Write-Host "PASS"
exit 0
