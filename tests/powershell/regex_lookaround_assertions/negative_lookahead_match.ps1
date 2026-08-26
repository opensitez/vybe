# vybe-test: powershell/regex_lookaround_assertions/negative_lookahead_match
$str = "cat and car and can"
$re = [regex]::new("ca(?!t)\w") # ca followed by something other than t
$matches = @($re.Matches($str) | ForEach-Object { $_.Value })
if ($matches.Count -ne 2 -or -not ($matches -contains "car") -or -not ($matches -contains "can")) {
    Write-Host "FAIL: Negative lookahead (?!t) failed"
    exit 1
}
Write-Host "PASS"
exit 0
