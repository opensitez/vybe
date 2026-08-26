# vybe-test: powershell/regex_named_capture_groups/system_text_regularexpressions_match_named_groups
$re = [regex]::new("(?<proto>https?)://(?<domain>[^/]+)")
$m = $re.Match("https://github.com/vybe")
if (-not $m.Success -or $m.Groups["proto"].Value -ne "https" -or $m.Groups["domain"].Value -ne "github.com") {
    Write-Host "FAIL: Regex object Match named groups failed"
    exit 1
}
Write-Host "PASS"
exit 0
