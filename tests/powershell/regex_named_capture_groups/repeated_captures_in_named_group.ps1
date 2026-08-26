# vybe-test: powershell/regex_named_capture_groups/repeated_captures_in_named_group
$re = [regex]::new("(?:(?<num>\d+),?)+")
$m = $re.Match("10,20,30")
$captures = @($m.Groups["num"].Captures)
if ($captures.Count -ne 3 -or $captures[0].Value -ne "10" -or $captures[2].Value -ne "30") {
    Write-Host "FAIL: Repeated captures in named group failed"
    exit 1
}
Write-Host "PASS"
exit 0
