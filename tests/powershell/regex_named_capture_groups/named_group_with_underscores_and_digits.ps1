# vybe-test: powershell/regex_named_capture_groups/named_group_with_underscores_and_digits
$str = "val: 42"
$matched = $str -match "val:\s*(?<my_group_1>\d+)"
if (-not $matched -or $Matches.my_group_1 -ne "42") {
    Write-Host "FAIL: Named group with underscores/digits failed"
    exit 1
}
Write-Host "PASS"
exit 0
